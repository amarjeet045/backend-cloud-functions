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
  sendResponse,
  hasSupportClaims,
} = require('../admin/utils');

const {
  code,
} = require('../admin/responses');


/**
 * Calls the resource related to an activity depending on the action
 * from the url.
 *
 * @param {Object} conn Contains Express' Request and Respone objects.
 */
const app = (conn) => {
  const action = require('url').parse(conn.req.url).path.split('/')[2];
  /** Can be used to verify in the activity flow to see if the request
   * is of type support.
   */
  conn.requester.isSupportRequest = false;

  if (conn.req.query.as === 'support') {
    conn.requester.isSupportRequest = true;
  }

  if (conn.requester.isSupportRequest
    && !hasSupportClaims(conn.requester.customClaims)) {
    sendResponse(
      conn,
      code.forbidden,
      'You do not have the permission to make support requests for activities.'
    );
    return;
  }

  if (action.startsWith('read')) {
    if (conn.req.method !== 'GET') {
      sendResponse(
        conn,
        code.methodNotAllowed,
        `${conn.req.method}`
        + ' is not allowed for the /read endpoint. Use GET.'
      );
      return;
    }

    const onRead = require('../firestore/activity/on-read');
    onRead(conn);
    return;
  }

  /** `/api/activities/create` can have a query string. */
  if (action.startsWith('create')) {
    if (conn.req.method !== 'POST') {
      sendResponse(
        conn,
        code.methodNotAllowed,
        `${conn.req.method} is not allowed for the /create`
        + ' endpoint. Use POST.'
      );
      return;
    }

    const onCreate = require('../firestore/activity/on-create');
    onCreate(conn);
    return;
  }

  if (action === 'comment') {
    if (conn.req.method !== 'POST') {
      sendResponse(
        conn,
        code.methodNotAllowed,
        `${conn.req.method} is not allowed for the /${action}`
        + ' endpoint. Use POST.'
      );
      return;
    }

    const onComment = require('../firestore/activity/on-comment');
    onComment(conn);
    return;
  }

  /** `/api/activities/update` can have a query string. */
  if (action.startsWith('update')) {
    if (conn.req.method !== 'PATCH') {
      sendResponse(
        conn,
        code.methodNotAllowed,
        `${conn.req.method} is not allowed for the /update endpoint. Use PATCH.`
        + ' Use PATCH.');
      return;
    }

    const onUpdate = require('../firestore/activity/on-update');
    onUpdate(conn);
    return;
  }

  /** `/api/activities/share` can have a query string. */
  if (action.startsWith('share')) {
    if (conn.req.method !== 'PATCH') {
      sendResponse(
        conn,
        code.methodNotAllowed,
        `${conn.req.method} is not allowed for the /share endpoint. Use PATCH.`
        + ' Use PATCH.'
      );
      return;
    }

    const onShare = require('../firestore/activity/on-share');
    onShare(conn);
    return;
  }

  if (action === 'change-status') {
    if (conn.req.method !== 'PATCH') {
      sendResponse(
        conn,
        code.methodNotAllowed,
        `${conn.req.method} is not allowed for the /${action}`
        + ' endpoint. Use PATCH.'
      );
      return;
    }

    const onStatusChange = require('../firestore/activity/on-change-status');
    onStatusChange(conn);
    return;
  }

  if (action === 'remove') {
    if (conn.req.method !== 'PATCH') {
      sendResponse(
        conn,
        code.methodNotAllowed,
        `${conn.req.method} is not allowed for the /${action}`
        + ' endpoint. Use PATCH.'
      );
      return;
    }

    const onRemove = require('../firestore/activity/on-remove');
    onRemove(conn);
    return;
  }

  sendResponse(
    conn,
    code.notFound,
    `No resource found at the path: ${(conn.req.url)}.`
  );
};


module.exports = app;
