/**
 * Copyright (c) 2018 GrowthFile
 *
 * Permission is hereby granted, free of charge, to any person obtaining a
 * copy of this software and associated documentation files (the "Software"),
 * to deal in the Software without restriction, including without limitation
 * the rights to use, copy, modify, merge, publish, distribute, sublicense,
 * and/or sell copies of the Software, and to permit persons to whom the
 * Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
 * THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
 * IN THE SOFTWARE.
 *
 */


'use strict';


const { rootCollections, getGeopointObject, db, } = require('../../admin/admin');

const { code, } = require('../../admin/responses');

const {
  handleError,
  sendResponse,
  getISO8601Date,
  isValidDate,
  isValidGeopoint,
  isE164PhoneNumber,
  isNonEmptyString,
} = require('../../admin/utils');



const commitBatch = (conn, locals) =>
  locals.batch.commit()
    .then(() => sendResponse(conn, code.noContent))
    .catch((error) => handleError(conn, error));


/**
 * Adds a doc in `/DailyActivities` collection in the path:
 * `/(office name)/(template name)` with the user's phone number,
 * timestamp of the request and the api used.
 *
 * @param {Object} conn Contains Express' Request and Response objects.
 * @param {Object} locals Object containing local data.
 * @returns {void}
 */
const updateDailyActivities = (conn, locals) => {
  const docId = getISO8601Date(locals.timestamp);

  locals.batch.set(rootCollections
    .dailyActivities
    .doc(docId)
    .collection('Logs')
    .doc(), {
      office: locals.activity.get('office'),
      timestamp: locals.timestamp,
      template: locals.activity.get('template'),
      phoneNumber: conn.requester.phoneNumber,
      url: conn.req.url,
      activityId: conn.req.body.activityId,
      geopoint: getGeopointObject(conn.req.body.geopoint),
    });

  commitBatch(conn, locals);
};


const updateActivityDoc = (conn, locals) => {
  locals.batch.set(rootCollections
    .activities
    .doc(conn.req.body.activityId), {
      timestamp: locals.timestamp,
    }, {
      merge: true,
    }
  );

  updateDailyActivities(conn, locals);
};


const setAddendumForUsersWithUid = (conn, locals) => {
  const promises = [];

  locals.assigneeArray.forEach((phoneNumber) => {
    promises.push(rootCollections.profiles.doc(phoneNumber).get());

    locals.batch.set(rootCollections
      .profiles
      .doc(phoneNumber)
      .collection('Activities')
      .doc(conn.req.body.activityId), {
        timestamp: locals.timestamp,
      }, {
        merge: true,
      }
    );
  });

  Promise
    .all(promises)
    .then((snapShot) => {
      snapShot.forEach((doc) => {
        /** `uid` is NOT `null` OR `undefined` */
        if (!doc.get('uid')) return;

        locals.batch.set(rootCollections
          .updates
          .doc(doc.get('uid'))
          .collection('Addendum')
          .doc(),
          locals.addendum
        );

      });

      updateActivityDoc(conn, locals);

      return;
    })
    .catch((error) => handleError(conn, error));
};


const unassignFromTheActivity = (conn, locals) => {
  let index;

  conn.req.body.remove.forEach((phoneNumber) => {
    if (!isE164PhoneNumber(phoneNumber)) return;

    /** Deleting from Assignees collection inside activity doc */
    locals.batch.delete(rootCollections
      .activities
      .doc(conn.req.body.activityId)
      .collection('Assignees')
      .doc(phoneNumber)
    );

    /** Deleting from Activities collection inside user Profile */
    locals.batch.delete(rootCollections
      .profiles
      .doc(phoneNumber)
      .collection('Activities')
      .doc(conn.req.body.activityId)
    );

    index = locals.assigneeArray.indexOf(phoneNumber);

    if (index > -1) {
      locals.assigneeArray.splice(index, 1);
    }
  });

  setAddendumForUsersWithUid(conn, locals);

  return;
};


const fetchTemplate = (conn, locals) => {
  const template = locals.activity.get('template');

  rootCollections
    .activityTemplates
    .doc(template)
    .get()
    .then((doc) => {
      locals.addendum = {
        activityId: conn.req.body.activityId,
        user: conn.requester.displayName || conn.requester.phoneNumber,
        comment: `${conn.requester.displayName || conn.requester.phoneNumber}`
          + ` updated ${doc.get('defaultTitle')}`,
        location: getGeopointObject(conn.req.body.geopoint),
        timestamp: locals.timestamp,
      };

      locals.template = doc;
      unassignFromTheActivity(conn, locals);

      return;
    })
    .catch((error) => handleError(conn, error));
};


const handleResult = (conn, result) => {
  if (!result[0].exists) {
    /** This case should probably never execute becase there is NO provision
     * for deleting an activity anywhere. AND, for reaching the `fetchDocs()`
     * function, the check for the existance of the activity has already
     * been performed in the `Profiles/(phoneNumber)/Activities(activity-id)`.
     */
    sendResponse(
      conn,
      code.conflict,
      `No activity found with the id: ${conn.req.body.activityId}.`
    );

    return;
  }

  /** Assignees collection in the `Activity/(doc-id)/Assignees` */
  if (result[1].size === 1) {
    /** An activity cannot exist with zero assignees. The person
     * last to stay cannot remove themselves.
     */
    sendResponse(
      conn,
      code.forbidden,
      `Cannot remove the last assignee of the activity.`
    );

    return;
  }

  /** Object for storing local data. */
  const locals = {};

  locals.batch = db.batch();

  /** Calling `new Date()` constructor multiple times is wasteful. */
  locals.timestamp = new Date(conn.req.body.timestamp);

  locals.activity = result[0];

  /** The `assigneeArray` is required to add addendum.
   * The `doc.id` is the phoneNumber of the assignee.
   */
  locals.assigneeArray = [];
  result[1].forEach((doc) => locals.assigneeArray.push(doc.id));

  fetchTemplate(conn, locals);
};


const fetchDocs = (conn) =>
  Promise
    .all([
      rootCollections
        .activities
        .doc(conn.req.body.activityId)
        .get(),
      rootCollections
        .activities
        .doc(conn.req.body.activityId)
        .collection('Assignees')
        .get(),
    ])
    .then((result) => handleResult(conn, result))
    .catch((error) => handleError(conn, error));


const verifyEditPermission = (conn) =>
  rootCollections
    .profiles
    .doc(conn.requester.phoneNumber)
    .collection('Activities')
    .doc(conn.req.body.activityId)
    .get()
    .then((doc) => {
      if (!doc.exists) {
        /** The activity does not exist in the system (OR probably
         * only for the user). */
        sendResponse(
          conn,
          code.notFound,
          `No activity found with the id: ${conn.req.body.activityId}.`
        );

        return;
      }

      if (!doc.get('canEdit')) {
        /** The `canEdit` flag is false so update is forbidden. */
        sendResponse(
          conn,
          code.forbidden,
          'You do not have the permission to edit this activity.'
        );

        return;
      }

      fetchDocs(conn);

      return;
    })
    .catch((error) => handleError(conn, error));


const isValidRequestBody = (body) =>
  isValidDate(body.timestamp)
  && isNonEmptyString(body.activityId)
  && Array.isArray(body.remove)
  && isValidGeopoint(body.geopoint);


module.exports = (conn) => {
  if (!isValidRequestBody(conn.req.body)) {
    sendResponse(
      conn,
      code.badRequest,
      'Invalid request body.'
      + ' Make sure to include the "activityId" (string), "timestamp" (long number),'
      + ' "remove" (array) and the "geopoint" (object) fields in the request body.'
    );

    return;
  }

  /** The support person doesn't need to be an assignee
   * of the activity to make changes.
   */
  if (conn.requester.isSupportRequest) {
    fetchDocs(conn);

    return;
  }

  verifyEditPermission(conn);
};
