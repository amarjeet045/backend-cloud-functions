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

const { handleCanEdit, isValidRequestBody, } = require('./helper');

const { code, } = require('../../admin/responses');

const {
  handleError,
  sendResponse,
  isE164PhoneNumber,
  logDailyActivities,
} = require('../../admin/utils');


/**
 * Updates the timestamp in the activity root document.
 *
 * @param {Object} conn Contains Express' Request and Response objects.
 * @param {Object} locals Object containing local data.
 * @returns {void}
 */
const updateActivityDoc = (conn, locals) => {
  locals.batch.set(rootCollections
    .activities
    .doc(conn.req.body.activityId), {
      timestamp: locals.timestamp,
    }, {
      merge: true,
    }
  );

  logDailyActivities(conn, locals, code.noContent);
};


/**
 * Updates the linked doc in the `docRef` field in the activity based on
 * the template name.
 *
 * @param {Object} conn Contains Express' Request and Response objects.
 * @param {Object} locals Object containing local data.
 * @returns {void}
 */
const updateLinkedDoc = (conn, locals) =>
  db
    .doc(locals.activity.get('docRef'))
    .get()
    .then((doc) => {
      const docData = doc.data();

      if (locals.activity.get('template') === 'subscription') {
        const includeArray = doc.get('include');
        const updatedIncludeArray = locals
          .validPhoneNumbers
          .concat(includeArray);

        docData.include = updatedIncludeArray;
      }

      if (locals.activity.get('template') === 'report') {
        const toArray = doc.get('to');
        const updatedToArray = locals
          .validPhoneNumbers
          .concat(toArray);

        docData.to = updatedToArray;
      }

      locals.batch.set(locals
        .activity
        .get('docRef'),
        docData
      );

      updateActivityDoc(conn, locals);

      return;
    })
    .catch((error) => handleError(conn, error));


/**
 * Handles the special case when the template name is 'report' or
 * 'subscription'.
 *
 * @param {Object} conn Contains Express' Request and Response objects.
 * @param {Object} locals Object containing local data.
 * @returns {void}
 */
const handleSpecialTemplates = (conn, locals) => {
  if (['subscription', 'report',]
    .indexOf(locals.activity.get('template')) > -1) {
    updateLinkedDoc(conn, locals);

    return;
  }

  updateActivityDoc(conn, locals);
};


/**
 * Adds the documents to batch for the users who have their `uid` populated
 * inside their profiles.
 *
 * @param {Object} conn Contains Express' Request and Response objects.
 * @param {Object} locals Object containing local data.
 * @returns {void}
 */
const setAddendumForUsersWithUid = (conn, locals) => {
  /** Assignee array can have duplicate elements. */
  const assigneeListWithUniques = Array.from(new Set(locals.assigneeArray));
  const promises = [];

  assigneeListWithUniques.forEach((phoneNumber) => {
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
        /** Create Profiles for the users who don't have a profile already. */
        if (!doc.exists) {
          /** The `doc.id` is the `phoneNumber` that doesn't exist */
          locals.batch.set(rootCollections
            .profiles
            .doc(doc.id), {
              uid: null,
            });
        }

        if (doc.exists && doc.get('uid')) {
          /** The `uid` is NOT `null` OR `undefined` */
          locals.batch.set(rootCollections
            .updates
            .doc(doc.get('uid'))
            .collection('Addendum')
            .doc(),
            locals.addendum
          );
        }
      });

      handleSpecialTemplates(conn, locals);

      return;
    })
    .catch((error) => handleError(conn, error));
};


/**
 * Adds addendum for all the assignees of the activity.
 *
 * @param {Object} conn Contains Express' Request and Response objects.
 * @param {Object} locals Object containing local data.
 * @returns {void}
 */
const addAddendumForAssignees = (conn, locals) => {
  let comment = `${conn.requester.phoneNumber} shared this activity with: `;

  locals.validPhoneNumbers = [];

  conn.req.body.share.forEach((phoneNumber) => {
    if (!isE164PhoneNumber(phoneNumber)) return;

    locals.validPhoneNumbers.push(phoneNumber);

    comment += `${phoneNumber}, `;

    /** The requester shouldn't be added to the activity assignee list
     * if the request is of `support` type.
     */
    if (phoneNumber === conn.requester.phoneNumber
      && conn.requester.isSupportRequest) return;

    /** Adding a doc with the id = phoneNumber in
     * `Activities/(activityId)/Assignees`
     * */
    locals.batch.set(rootCollections
      .activities
      .doc(conn.req.body.activityId)
      .collection('Assignees')
      .doc(phoneNumber), {
        canEdit: handleCanEdit(
          locals,
          phoneNumber,
          conn.requester.phoneNumber
        ),
      }, {
        merge: true,
      }
    );

    /** Adding a doc with the id = activityId inside
     *  Profiles/(phoneNumber)/Activities/(activityId)
     * */
    locals.batch.set(rootCollections
      .profiles
      .doc(phoneNumber)
      .collection('Activities')
      .doc(conn.req.body.activityId), {
        canEdit: handleCanEdit(
          locals,
          phoneNumber,
          conn.requester.phoneNumber
        ),
        timestamp: locals.timestamp,
      }, {
        merge: true,
      }
    );

    locals.assigneeArray.push(phoneNumber);
  });

  locals.addendum.comment = comment.trim();

  setAddendumForUsersWithUid(conn, locals);
};


const handleResult = (conn, result) => {
  if (!result[0].exists) {
    /** This case should probably never execute because there is NO provision
     * for deleting an activity anywhere. AND, for reaching the fetchDocs()
     * function, the check for the existence of the activity has already
     * been performed in the User's profile.
     */
    sendResponse(
      conn,
      code.conflict,
      `No activity found with the id: ${conn.req.body.activityId}.`
    );

    return;
  }

  const locals = {};
  locals.batch = db.batch();

  /** Calling `new Date()` constructor multiple times is wasteful. */
  locals.timestamp = new Date(conn.req.body.timestamp);

  locals.activity = result[0];

  locals.canEditRule = result[0].get('canEditRule');

  /** The assigneeArray is required to add addendum.
   * The `doc.id` is the phoneNumber of the assignee.
   */
  locals.assigneeArray = [];
  result[1].forEach((doc) => locals.assigneeArray.push(doc.id));

  /** Comment field will be added later. */
  locals.addendum = {
    activityId: conn.req.body.activityId,
    user: conn.requester.phoneNumber,
    location: getGeopointObject(conn.req.body.geopoint),
    timestamp: locals.timestamp,
  };

  addAddendumForAssignees(conn, locals);
};


/**
 * Fetches the activity doc, along with all the `assignees` of the activity
 * using the `activityId` from the `request body`.
 *
 * @param {Object} conn Contains Express' Request and Response objects.
 * @returns {void}
 */
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


/**
 * Checks if the requester has the permission to perform an update
 * to this activity. For this to happen, the `canEdit` flag is checked.
 *
 * @param {Object} conn Contains Express' Request and Response objects.
 * @returns {void}
 */
const verifyEditPermission = (conn) =>
  rootCollections
    .profiles
    .doc(conn.requester.phoneNumber)
    .collection('Activities')
    .doc(conn.req.body.activityId)
    .get()
    .then((doc) => {
      if (!doc.exists) {
        /** The activity doesn't exist for the user */
        sendResponse(
          conn,
          code.notFound,
          `An activity with the id: ${conn.req.body.activityId} doesn't exist.`
        );

        return;
      }

      if (!doc.get('canEdit')) {
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


module.exports = (conn) => {
  const result = isValidRequestBody(conn.req.body, 'share');

  if (!result.isValidBody) {
    sendResponse(
      conn,
      code.badRequest,
      result.message
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
