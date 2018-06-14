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


const {
  rootCollections,
  getGeopointObject,
  db,
} = require('../../admin/admin');

const {
  handleError,
  sendResponse,
} = require('../../admin/utils');

const {
  isValidDate,
  isValidString,
  isValidLocation,
} = require('./helper');

const {
  code,
} = require('../../admin/responses');

const {
  activities,
  updates,
  profiles,
  dailyActivities,
} = rootCollections;


/**
 * Commits the batch to the Firestore and send a response to the client
 * about the result.
 *
 * @param {Object} conn Object with Express Request and Response Objects.
 */
const commitBatch = (conn) => conn.batch.commit()
  .then(() => sendResponse(
    conn,
    code.created,
    'The comment was successfully added to the activity.'
  )).catch((error) => handleError(conn, error));


/**
 * Adds a doc in `/DailyActivities` collection in the path:
 * `/(office name)/(template name)` with the user's phone number,
 * timestamp of the request and the api used.
 *
* @param {Object} conn Contains Express' Request and Response objects.
 */
const updateDailyActivities = (conn) => {
  conn.batch.set(dailyActivities.doc(new Date().toDateString())
    .collection(conn.data.activity.get('office'))
    .doc(conn.data.activity.get('template')), {
      phoneNumber: conn.requester.phoneNumber,
      url: conn.req.url,
      timestamp: new Date(),
      activityId: conn.req.body.activityId,
    });

  commitBatch(conn);
};


/**
* Creates a doc inside `/Profiles/(phoneNumber)/Map` for tracking location
* history of the user.
*
* @param {Object} conn Contains Express' Request and Response objects.
*/
const logLocation = (conn) => {
  conn.batch.set(profiles.doc(conn.requester.phoneNumber).collection('Map')
    .doc(), {
      geopoint: getGeopointObject(conn.req.body.geopoint),
      timestamp: new Date(conn.req.body.timestamp),
      office: conn.data.activity.get('office'),
      template: conn.data.activity.get('template'),
    });

  updateDailyActivities(conn);
};


/**
 * Adds addendum doc for each assignee of the activity for which the comment
 * is being created.
 *
 * @param {Object} conn Object with Express Request and Response Objects.
 */
const setAddendumForAssignees = (conn) => {
  Promise.all(conn.assigneeDocPromises).then((snapShots) => {
    snapShots.forEach((doc) => {
      /** `uid` shouldn't be `null` OR `undefined` */
      if (doc.exists && doc.get('uid')) {
        conn.batch.set(updates.doc(doc.get('uid')).collection('Addendum')
          .doc(), {
            activityId: conn.req.body.activityId,
            user: conn.requester.displayName || conn.requester.phoneNumber,
            comment: conn.req.body.comment,
            location: getGeopointObject(conn.req.body.geopoint),
            timestamp: new Date(conn.req.body.timestamp),
          });
      }
    });

    logLocation(conn);
    return;
  }).catch((error) => handleError(conn, error));
};


/**
 * Fetches all the docs from '/Assignees' subcollection in the activity
 * and creates a list of profiles for which the Addendum are to be written.
 *
 * @param {Object} conn Object with Express Request and Response Objects.
 */
const constructActivityAssigneesPromises = (conn) => {
  conn.assigneeDocPromises = [];

  activities.doc(conn.req.body.activityId).collection('Assignees').get()
    .then((snapShot) => {
      snapShot.forEach((doc) =>
        conn.assigneeDocPromises.push(profiles.doc(doc.id).get()));

      conn.batch = db.batch();

      setAddendumForAssignees(conn);
      return;
    }).catch((error) => handleError(conn, error));
};


/**
 * Checks whether the user is an assignee to an activity which they
 * have sent a request to add a comment to.
 *
 * @param {Object} conn Object with Express Request and Response Objects.
 */
const checkCommentPermission = (conn) => {
  if (!conn.data.profileActivityDoc.exists) {
    sendResponse(
      conn,
      code.notFound,
      `No acivity found with the id: ${conn.req.body.activityId}`
    );
    return;
  }

  constructActivityAssigneesPromises(conn);
};


const fetchDocs = (conn) => {
  Promise.all([
    profiles.doc(conn.requester.phoneNumber).collection('Activities')
      .doc(conn.req.body.activityId).get(),
    activities.doc(conn.req.body.activityId).get(),
  ]).then((docsArray) => {
    conn.data = {};
    conn.data.profileActivityDoc = docsArray[0];
    conn.data.activity = docsArray[1];

    checkCommentPermission(conn);
    return;
  }).catch((error) => handleError(conn, error));
};


const app = (conn) => {
  if (isValidDate(conn.req.body.timestamp)
    && isValidLocation(conn.req.body.geopoint)
    && isValidString(conn.req.body.activityId)
    && typeof conn.req.body.comment === 'string') {
    fetchDocs(conn);
    return;
  }

  sendResponse(
    conn,
    code.badRequest,
    'The request body does not have all the necessary fields with proper'
    + ' values. Please make sure that the timestamp, activityId  and the '
    + 'geopoint are included in the request body.'
  );
};


module.exports = app;
