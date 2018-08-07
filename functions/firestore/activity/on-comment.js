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


const {
  rootCollections,
  getGeopointObject,
  db,
  serverTimestamp,
} = require('../../admin/admin');

const { isValidRequestBody, } = require('./helper');

const { code, } = require('../../admin/responses');

const {
  handleError,
  sendResponse,
} = require('../../admin/utils');


/**
 * Creates a document in the path: `/AddendumObjects/(auto-id)`.
 * This will trigger an auto triggering cloud function which will
 * copy this addendum to ever assignee's `/Updates/(uid)/Addendum(auto-id)`
 * doc.
 *
 * @param {Object} conn Object with Express Request and Response Objects.
 * @param {Object} locals Object containing local data.
 * @returns {void}
 */
const createAddendumDoc = (conn, locals) => {
  const docRef = rootCollections
    .offices
    .doc(locals.activity.get('officeId'))
    .collection('Addendum')
    .doc();

  locals.batch.set(docRef, {
    activityId: conn.req.body.activityId,
    user: conn.requester.phoneNumber,
    comment: conn.req.body.comment,
    location: getGeopointObject(conn.req.body.geopoint),
    userDeviceTimestamp: new Date(conn.req.body.timestamp),
    timestamp: serverTimestamp,
  });

  /** ENDS the response. */
  locals
    .batch
    .commit()
    .then(() => sendResponse(conn, code.noContent))
    .catch((error) => handleError(conn, error));
};


/**
 * Checks whether the user is an assignee to an `activity` which they
 * have sent a request to add a comment to.
 *
 * @param {Object} conn Object with Express Request and Response Objects.
 * @param {Object} locals Object containing local data.
 * @returns {void}
 */
const checkCommentPermission = (conn, locals) => {
  if (!locals.profileActivityDoc.exists) {
    sendResponse(
      conn,
      code.badRequest,
      `No activity found with the id: ${conn.req.body.activityId}.`
    );

    return;
  }

  createAddendumDoc(conn, locals);

  return;
};


/**
 * Fetches the `activity` doc from user's `Subscription` and the
 * `Activities` collection.
 *
 * @param {Object} conn Contains Express Request and Response Objects.
 * @returns {void}
 */
const fetchDocs = (conn) =>
  Promise
    .all([
      rootCollections
        .profiles
        .doc(conn.requester.phoneNumber)
        .collection('Activities')
        .doc(conn.req.body.activityId)
        .get(),
      rootCollections
        .activities
        .doc(conn.req.body.activityId)
        .get(),
    ])
    .then((result) => {
      const locals = {
        batch: db.batch(),
        profileActivityDoc: result[0],
        activity: result[1],
      };

      checkCommentPermission(conn, locals);

      return;
    })
    .catch((error) => handleError(conn, error));


module.exports = (conn) => {
  const result = isValidRequestBody(conn.req.body, 'comment');

  if (!result.isValidBody) {
    sendResponse(
      conn,
      code.badRequest,
      result.message
    );

    return;
  }

  fetchDocs(conn);
};
