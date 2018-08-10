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
  db,
  users,
  rootCollections,
  serverTimestamp,
  getGeopointObject,
} = require('../../admin/admin');

const { code, } = require('../../admin/responses');

const {
  validateVenues,
  getCanEditValue,
  filterAttachment,
  validateSchedules,
  isValidRequestBody,
} = require('./helper');

const {
  handleError,
  sendResponse,
} = require('../../admin/utils');


const createDocsWithBatch = (conn, locals) => {
  locals.objects.allPhoneNumbers
    .forEach((phoneNumber) => {
      const isRequester = phoneNumber === conn.requester.phoneNumber;

      /**
       * Support requests won't add the creator to the
       * activity assignee list.
       */
      if (isRequester && conn.requester.isSupportRequest) return;

      locals.batch.set(locals.docs.activityRef
        .collection('Assignees')
        .doc(phoneNumber), {
          activityId: locals.docs.activityRef.id,
          canEdit: getCanEditValue(locals, phoneNumber),
        });
    });

  locals.batch.set(locals.docs.activityRef, {
    docRef: locals.docs.docRef,
    venue: conn.req.body.venue,
    timestamp: serverTimestamp,
    office: conn.req.body.office,
    template: conn.req.body.template,
    schedule: conn.req.body.schedule,
    status: locals.static.statusOnCreate,
    attachment: conn.req.body.attachment || {},
    canEditRule: locals.static.canEditRule,
    activityName: conn.req.body.activityName || '',
    officeId: rootCollections.offices.doc(locals.static.officeId).id,
  });

  locals.batch.set(rootCollections
    .offices
    .doc(locals.static.officeId)
    .collection('Addendum')
    .doc(), {
      remove: null,
      action: 'create',
      updatedPhoneNumber: null,
      timestamp: serverTimestamp,
      user: conn.requester.phoneNumber,
      activityId: locals.static.activityId,
      template: conn.req.body.template,
      share: conn.req.body.share || [],
      location: getGeopointObject(conn.req.body.geopoint),
      userDeviceTimestamp: new Date(conn.req.body.timestamp),
      updatedFields: [],
    });

  /** ENDS the response. */
  locals.batch.commit()
    .then(() => sendResponse(
      conn,
      code.created,
      'The activity was successfully created.'
    ))
    .catch((error) => handleError(conn, error));
};


const handleAssignees = (conn, locals) => {
  const promises = [];

  locals
    .objects
    .allPhoneNumbers
    .forEach((phoneNumber) => {
      const isRequester = phoneNumber === conn.requester.phoneNumber;

      /**
       * Support requests won't add the creator to the
       * activity assignee list.
       */
      if (isRequester && conn.requester.isSupportRequest) return;

      locals.objects.permissions[phoneNumber] = {
        isAdmin: false,
        isEmployee: false,
        isCreator: isRequester,
      };

      /**
       * No docs will exist if the template is `office`
       * since this template itself is used to create
       * the office. No use of adding promises to the array.
       */
      if (conn.req.body.template === 'office') return;

      const officeId = locals.static.officeId;

      promises.push(rootCollections
        .offices.doc(officeId)
        .collection('Activities')
        .where('attachment.Phone Number.value', '==', phoneNumber)
        .where('template', '==', 'admin')
        .limit(1)
        .get()
      );

      promises.push(rootCollections
        .offices.doc(officeId)
        .collection('Activities')
        .where('attachment.Phone Number.value', '==', phoneNumber)
        .where('template', '==', 'employee')
        .limit(1)
        .get()
      );
    });

  if (promises.length === 0) {
    createDocsWithBatch(conn, locals);

    return;
  }

  Promise
    .all(promises)
    .then((snapShots) => {
      snapShots.forEach((snapShot) => {
        if (snapShot.empty) return;

        const doc = snapShot.docs[0];
        const template = doc.get('template');
        const phoneNumber = doc.get('attachment.Phone Number.value');

        /** The person can either be an `employee` or an `admin`. */
        if (template === 'admin') {
          locals.objects.permissions[phoneNumber].isAdmin = true;

          return;
        }

        locals.objects.permissions[phoneNumber].isEmployee = true;
      });

      createDocsWithBatch(conn, locals);

      return;
    })
    .catch((error) => handleError(conn, error));
};


const handleExtra = (conn, locals) => {
  const scheduleNames = locals.objects.schedule;
  const scheduleValid = validateSchedules(conn.req.body, scheduleNames);

  if (!scheduleValid.isValid) {
    sendResponse(conn, code.badRequest, scheduleValid.message);

    return;
  }

  const venueDescriptors = locals.objects.venue;
  const venueValid = validateVenues(conn.req.body, venueDescriptors);

  if (!venueValid.isValid) {
    sendResponse(conn, code.badRequest, venueValid.message);

    return;
  }

  const attachmentValid = filterAttachment(conn.req.body, locals);

  if (!attachmentValid.isValid) {
    sendResponse(conn, code.badRequest, attachmentValid.message);

    return;
  }

  attachmentValid.phoneNumbers
    .forEach((phoneNumber) => {
      if (phoneNumber === '') return;

      locals.objects.allPhoneNumbers.add(phoneNumber);
    });

  if (!attachmentValid.promise) {
    handleAssignees(conn, locals);

    return;
  }

  attachmentValid
    .promise
    .then((snapShot) => {
      if (!snapShot.empty) {
        const value = conn.req.body.attachment.Name.value;
        const type = conn.req.body.attachment.Name.type;
        const message = `'${value}' already exists in the office`
          + ` '${conn.req.body.office}' with the template '${type}'.`;

        sendResponse(conn, code.conflict, message);

        return;
      }

      handleAssignees(conn, locals);

      return;
    })
    .catch((error) => handleError(conn, error));
};


const createLocals = (conn, result) => {
  const activityRef = rootCollections.activities.doc();

  /**
   * Temporary object in memory to store all data during the function
   * instance.
   */
  const locals = {
    batch: db.batch(),
    /**
     * Stores all the static data during the function instance.
     */
    static: {
      /** Storing this here to be consistent with other functions. */
      activityId: activityRef.id,
      /**
       * A fallback case when the template is `office` so the
       * activity is used to create the office. This value will
       * updated accordingly at appropriate time after checking
       * the template name from the request body.
       */
      officeId: activityRef.id,
      /**
       * A fallback in cases when the subscription doc is not found
       * during the `support` requests.
       */
      include: [],
      canEditRule: null,
      statusOnCreate: null,
    },
    /**
     * For storing all object types (e.g, schedule, venue, attachment)
     *  for the function instance.
     */
    objects: {
      /**
       * Using a `Set()` to avoid duplication of phone numbers.
       */
      allPhoneNumbers: new Set(),
      /** Stores the phoneNumber and it's permission to see
       * if it is an `admin` of the office, or an `employee`.
       */
      permissions: {},
      schedule: [],
      venue: [],
      attachment: {},
    },
    /**
     * Stores all the document references for the function instance.
     */
    docs: {
      activityRef,
      /**
       * Points to the document which this activity was used to create.
       * This either points to an `office` doc, or an activity doc
       * which is a child to that `office`.
       *
       * @description The `docRef` is the same as the `activityId`
       * for the case when the template name is `office`. For any
       * other case, like (e.g., template name === 'employee'), this
       * value will be updated to point to a document inside
       * a sub-collection in the path
       * `Offices/(officeId)/Activities/(activityId)`.
       */
      docRef: rootCollections.offices.doc(activityRef.id),
    },
  };

  if (!conn.requester.isSupportRequest) {
    locals.objects.allPhoneNumbers.add(conn.requester.phoneNumber);
  }

  const [
    templateQueryResult,
    subscriptionQueryResult,
    officeQueryResult,
  ] = result;

  if (templateQueryResult.empty) {
    sendResponse(
      conn,
      code.badRequest,
      `Template '${conn.req.body.template}' not found.`
    );

    return;
  }

  locals.objects.schedule = templateQueryResult.docs[0].get('schedule');
  locals.objects.venue = templateQueryResult.docs[0].get('venue');
  locals.objects.attachment = templateQueryResult.docs[0].get('attachment');

  locals.static.canEditRule = templateQueryResult.docs[0].get('canEditRule');
  locals.static.statusOnCreate = templateQueryResult.docs[0].get('statusOnCreate');
  /** Used by the filterAttachment function to query the
   * `Office/(officeId)/Activities` collection by using the
   * attachment.Name.value. */
  locals.static.template = templateQueryResult.docs[0].get('name');

  if (subscriptionQueryResult.empty && !conn.requester.isSupportRequest) {
    sendResponse(
      conn,
      code.forbidden,
      `No subscription found for the template: '${conn.req.body.template}'`
      + ` with the office '${conn.req.body.office}'.`
    );

    return;
  }

  if (!subscriptionQueryResult.empty) {
    if (subscriptionQueryResult.docs[0].get('status') === 'CANCELLED') {
      sendResponse(
        conn,
        code.forbidden,
        `Your subscription to the template '${conn.req.body.template}'`
        + ` is 'CANCELLED'. Cannot create an activity.`
      );

      return;
    }

    /**
   * Default assignees for all the activities that the user
   * creates using the subscription mentioned in the request body.
   */
    subscriptionQueryResult.docs[0].get('include')
      .forEach(
        (phoneNumber) => locals.objects.allPhoneNumbers.add(phoneNumber)
      );
  }

  if (!officeQueryResult.empty) {
    if (conn.req.body.template === 'office') {
      sendResponse(
        conn,
        code.conflict,
        `The office '${conn.req.body.office}' already exists.`
      );

      return;
    }

    if (officeQueryResult.docs[0].get('status') === 'CANCELLED') {
      sendResponse(
        conn,
        code.forbidden,
        `The office status is 'CANCELLED'. Cannot create an activity.`
      );

      return;
    }

    const officeId = officeQueryResult.docs[0].id;

    locals.static.officeId = officeId;
    locals.docs.docRef =
      rootCollections
        .offices
        .doc(officeId)
        .collection('Activities')
        .doc(locals.static.activityId);
  }

  if (officeQueryResult.empty) {
    if (conn.req.body.office !== conn.req.body.attachment.Name.value) {
      sendResponse(
        conn,
        code.conflict,
        `The office name in the 'attachment.Name.value' and the`
        + ` 'office' field should be the same.`
      );

      return;
    }
  }

  if (conn.req.body.hasOwnProperty('share')) {
    if (conn.req.body.share.length === 0
      && locals.static.include.length === 0) {
      sendResponse(
        conn,
        code.conflict,
        `Cannot create an activity without any assignees.`
      );

      return;
    }

    conn.req.body.share
      .forEach((phoneNumber) =>
        locals.objects.allPhoneNumbers.add(phoneNumber));
  }

  handleExtra(conn, locals);
};


const fetchDocs = (conn) => {
  Promise
    .all([
      rootCollections
        .activityTemplates
        .where('name', '==', conn.req.body.template)
        .limit(1)
        .get(),
      rootCollections
        .profiles
        .doc(conn.requester.phoneNumber)
        .collection('Subscriptions')
        .where('office', '==', conn.req.body.office)
        .where('template', '==', conn.req.body.template)
        .limit(1)
        .get(),
      rootCollections
        .offices
        .where('attachment.Name.value', '==', conn.req.body.office)
        .limit(1)
        .get(),
    ])
    .then((result) => createLocals(conn, result))
    .catch((error) => handleError(conn, error));
};


module.exports = (conn) => {
  const bodyResult = isValidRequestBody(conn.req.body, 'create');

  if (!bodyResult.isValid) {
    sendResponse(conn, code.badRequest, bodyResult.message);

    return;
  }

  if (conn.req.body.template !== 'admin') {
    fetchDocs(conn);

    return;
  }

  if (!conn.req.body.attachment.hasOwnProperty('Phone Number')) {
    sendResponse(
      conn,
      code.badRequest,
      `The 'Phone Number' field is missing from the attachment object.`
    );

    return;
  }

  /**
   * Phone number of the user who's being given the `admin` custom
   * claims with the `admin` template.
   */
  if (!conn.req.body.attachment['Phone Number'].hasOwnProperty('value')) {
    sendResponse(
      conn,
      code.badRequest,
      `The 'value' field is missing from the 'Phone Number'`
      + ` object in 'attachment'.`
    );

    return;
  }

  const isE164PhoneNumber = require('../../admin/utils').isE164PhoneNumber;

  if (!isE164PhoneNumber(conn.req.body.attachment['Phone Number'].value)) {
    sendResponse(
      conn,
      code.badRequest,
      `The 'Phone Number'.value field in the 'attachment' does not have`
      + ` a valid phone number. Cannot create an admin for the`
      + `office ${conn.req.body.office}.`
    );

    return;
  }

  const phoneNumber = conn.req.body.attachment['Phone Number'].value;

  users
    .getUserByPhoneNumber(phoneNumber)
    .then((userRecord) => {
      const phoneNumber = Object.keys(userRecord)[0];
      const record = userRecord[`${phoneNumber}`];

      if (!record.hasOwnProperty('uid')) {
        sendResponse(
          conn,
          code.forbidden,
          `No user found with the phone number: '${phoneNumber}'.`
          + ` Granting admin rights is not possible.`
        );

        return;
      }

      fetchDocs(conn);

      return;
    })
    .catch((error) => handleError(conn, error));
};
