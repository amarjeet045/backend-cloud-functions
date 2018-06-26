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
} = require('../../admin/admin');

const {
  handleError,
  sendResponse,
} = require('../../admin/utils');

const {
  isValidDate,
  isValidString,
  isValidLocation,
  filterSchedules,
  filterVenues,
} = require('./helper');

const {
  code,
} = require('../../admin/responses');

const {
  activities,
  profiles,
  updates,
  activityTemplates,
  dailyActivities,
} = rootCollections;


/**
 * Commits the batch to the DB.
 *
 * @param {Object} conn Contains Express' Request and Respone objects.
 * @returns {Promise} Batch Object.
 */
const commitBatch = (conn) => conn.batch.commit()
  .then((data) => sendResponse(conn, code.noContent))
  .catch((error) => handleError(conn, error));


/**
 * Adds a doc in `/DailyActivities` collection in the path:
 * `/(office name)/(template name)` with the user's phone number,
 * timestamp of the request and the api used.
 *
 * @param {Object} conn Contains Express' Request and Response objects.
 * @returns {void}
 */
const updateDailyActivities = (conn) => {
  const date = new Date();

  const hour = date.getHours();
  const minutes = date.getMinutes();
  const seconds = date.getSeconds();

  const office = conn.data.activity.get('office');
  const template = conn.data.activity.get('template');

  const dailyActivitiesDoc =
    dailyActivities
      .doc(date.toDateString())
      .collection(office)
      .doc(template);

  const data = {
    [`${hour}h:${minutes}m:${seconds}s`]: {
      phoneNumber: conn.requester.phoneNumber,
      url: conn.req.url,
      activityId: conn.req.body.activityId,
    },
  };

  conn.batch.set(dailyActivitiesDoc, data);

  commitBatch(conn);
};


/**
 * Creates a doc inside `/Profiles/(phoneNumber)/Map` for tracking location
 * history of the user.
 *
 * @param {Object} conn Contains Express' Request and Response objects.
 * @returns {void}
 */
const logLocation = (conn) => {
  const locationDoc =
    profiles
      .doc(conn.requester.phoneNumber)
      .collection('Map')
      .doc();

  const data = {
    activityId: conn.req.body.activityId,
    geopoint: getGeopointObject(conn.req.body.geopoint),
    timestamp: new Date(conn.req.body.timestamp),
    office: conn.data.activity.get('office'),
    template: conn.data.activity.get('template'),
  };

  conn.batch.set(locationDoc, data);

  updateDailyActivities(conn);
};


/**
 * Updates the activity root and adds the data to the batch.
 *
 * @param {Object} conn Contains Express' Request and Respone objects.
 * @param {Object} update Fields for the activity root object.
 * @returns {void}
 */
const updateActivityDoc = (conn, update) => {
  if (conn.req.body.hasOwnProperty('title')
    && isValidString(conn.req.body.title)) {
    update.title = conn.req.body.title;
  }

  if (conn.req.body.hasOwnProperty('description')
    && isValidString(conn.req.body.dailyActivitiesDoc)) {
    update.description = conn.req.body.description;
  }

  if (conn.req.body.hasOwnProperty('schedule')) {
    update.schedule = filterSchedules(
      conn.req.body.schedule,
      /** The schedule is an array of objects in Firestore.
       * For comparing the venues, we only need a single object.
      */
      conn.data.activity.get('schedule')[0]
    );
  }

  if (conn.req.body.hasOwnProperty('venue')) {
    update.venue = filterVenues(
      conn.req.body.venue,
      /** The venue is an array of objects in Firestore. For
       * comparing the venues, we only need a single object.
       */
      conn.data.activity.get('venue')[0]
    );
  }

  update.timestamp = new Date(conn.req.body.timestamp);

  /** Imeplementing the `handleAttachment()` method will make this work. */
  if (conn.hasOwnProperty('docRef')) {
    /**
     * The `docRef` is not `undefined` only when a document is updated during
     * the update operation.
     */
    updates.docRef = conn.docRef;
  }

  conn.batch.set(
    activities.doc(conn.req.body.activityId),
    update, {
      /** The activity doc *will* have some of these fields by default. */
      merge: true,
    }
  );

  logLocation(conn);
};



/**
 * Manages the attachment object.
 *
 * @param {Object} conn Contains Express' Request and Respone objects.
 * @param {Object} update Fields for the activity root object.
 * @returns {void}
 */
const handleAttachment = (conn, update) => {
  if (!conn.req.body.hasOwnProperty('attachment')) {
    updateActivityDoc(conn, update);
    return;
  }

  /** Do stuff */
  updateActivityDoc(conn, update);
};



/**
 * Adds addendum data for all the assignees in the activity.
 *
 * @param {Object} conn Contains Express' Request and Respone objects.
 * @returns {void}
 */
const addAddendumForAssignees = (conn) => {
  conn.data.assigneesPhoneNumbersArray.forEach((phoneNumber) => {
    conn.batch.set(
      profiles
        .doc(phoneNumber)
        .collection('Activities')
        .doc(conn.req.body.activityId), {
        timestamp: new Date(conn.req.body.timestamp),
      }, {
        merge: true,
      }
    );
  });

  Promise.all(conn.data.assigneesArray).then((snapShot) => {
    snapShot.forEach((doc) => {
      if (!doc.get('uid')) return;

      /** Users without `uid` are the ones who don't have
       * signed up. Addemdum is added only for the users who
       * have an account in auth.
       */
      conn.batch.set(
        updates
          .doc(doc.get('uid'))
          .collection('Addendum')
          .doc(),
        conn.data.addendum
      );
    });

    /** Stores the objects that are to be updated in the activity root. */
    const update = {};

    handleAttachment(conn, update);
    return;
  }).catch((error) => handleError(conn, error));
};


/**
 * Gets the template from the activity root.
 *
 * @param {Object} conn Contains Express' Request and Respone objects.
 * @returns {void}
 */
const fetchTemplate = (conn) => {
  activityTemplates
    .doc(conn.data.activity.get('template'))
    .get()
    .then((doc) => {
      conn.data.addendum = {
        activityId: conn.req.body.activityId,
        user: conn.requester.displayName || conn.requester.phoneNumber,
        comment: `${conn.requester.displayName || conn.requester.phoneNumber}`
          + ` updated ${doc.get('defaultTitle')}`,
        location: getGeopointObject(conn.req.body.geopoint),
        timestamp: new Date(conn.req.body.timestamp),
      };

      conn.data.template = doc;
      addAddendumForAssignees(conn);

      return;
    })
    .catch((error) => handleError(conn, error));
};


/**
 * Fetches the activity, and its assignees from the DB.
 *
 * @param {Object} conn Contains Express' Request and Respone objects.
 * @returns {void}
 */
const fetchDocs = (conn) => {
  Promise.all([
    activities
      .doc(conn.req.body.activityId)
      .get(),
    activities
      .doc(conn.req.body.activityId)
      .collection('Assignees')
      .get(),
  ]).then((result) => {
    if (!result[0].exists) {
      /** This case should probably never execute becase there is provision
       * for deleting an activity anywhere. AND, for reaching the fetchDocs()
       * function, the check for the existance of the activity has already
       * been performed in the User's profile.
       */
      sendResponse(
        conn,
        code.conflict,
        `There is no activity with the id: ${conn.req.body.activityId}`
      );
      return;
    }

    conn.batch = db.batch();

    conn.data = {};
    conn.data.activity = result[0];

    conn.data.assigneesArray = [];
    conn.data.assigneesPhoneNumbersArray = [];

    result[1].forEach((doc) => {
      /** The assigneesArray is required to add addendum. */
      conn.data.assigneesArray.push(profiles.doc(doc.id).get());
      conn.data.assigneesPhoneNumbersArray.push(doc.id);
    });

    fetchTemplate(conn);
    return;
  }).catch((error) => handleError(conn, error));
};


/**
 * Checks if the user has permission to update the activity data.
 *
 * @param {Object} conn Contains Express' Request and Respone objects.
 * @returns {void}
 */
const verifyEditPermission = (conn) => {
  profiles
    .doc(conn.requester.phoneNumber)
    .collection('Activities')
    .doc(conn.req.body.activityId)
    .get()
    .then((doc) => {
      if (!doc.exists) {
        /** The activity doesn't exist for the user */
        sendResponse(
          conn,
          conn.forbidden,
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
};


const isValidRequestBody = (conn) => {
  return isValidDate(conn.req.body.timestamp)
    && isValidString(conn.req.body.activityId)
    && isValidLocation(conn.req.body.geopoint);
};


const app = (conn) => {
  if (!isValidRequestBody(conn)) {
    sendResponse(
      conn,
      code.badRequest,
      `The request body is invalid. Make sure that the activityId, timestamp`
      + ` and the geopoint are included.`
    );
    return;
  }

  if (conn.requester.isSupportRequest) {
    fetchDocs(conn);
    return;
  }

  verifyEditPermission(conn);
};


module.exports = app;
