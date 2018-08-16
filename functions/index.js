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


const functions = require('firebase-functions');

const authOnCreate = functions
  .auth
  .user()
  .onCreate(require('./auth/on-create'));

const authOnDelete = functions
  .auth
  .user()
  .onDelete(require('./auth/on-delete'));

const api = functions
  .https
  .onRequest(require('./server/server'));

const addendumHandler = functions
  .firestore
  .document('Offices/{officeId}/Addendum/{docId}')
  .onCreate(require('./firestore/addendum/index'));

const assigneeHandler = functions
  .firestore
  .document('Activities/{activityId}/Assignees/{phoneNumber}')
  .onWrite(require('./firestore/assignees/index'));

const activityHandler = functions
  .firestore
  .document('/Activities/{activityId}')
  .onWrite(require('./firestore/activity/on-write'));

const phoneNumberUpdateHandler = functions
  .firestore
  .document('PhoneNumberUpdates/docId')
  .onCreate(require('./firestore/profiles/index'));

const instantMail = functions
  .firestore
  .document('Instant/{docId}')
  .onCreate(require('./firestore/instant/index'));


module.exports = {
  api,
  instantMail,
  authOnCreate,
  authOnDelete,
  addendumHandler,
  assigneeHandler,
  activityHandler,
  phoneNumberUpdateHandler,
};
