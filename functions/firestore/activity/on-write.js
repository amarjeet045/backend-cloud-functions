/**
 * Copyright (c) 2020 GrowthFile
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

const { db, rootCollections, auth } = require('../../admin/admin');
const {
  vowels,
  httpsActions,
  dateFormats,
  reportNames,
  addendumTypes,
  subcollectionNames,
  msEndpoints,
  msRequestTypes,
} = require('../../admin/constants');

const {
  slugify,
  getAuth,
  getLatLngString,
  adjustedGeopoint,
  isNonEmptyString,
  getNumbersbetween,
  getDefaultAttendanceObject,
  getDistanceFromDistanceMatrix,
  growthfileMsRequester,
  enumerateDaysBetweenDates,
} = require('../../admin/utils');
const { toMapsUrl, getStatusForDay } = require('../recipients/report-utils');
const {
  haversineDistance,
  createAutoSubscription,
} = require('../activity/helper');
const connectionToExternal = require('../../external');
const env = require('../../admin/env');
const admin = require('firebase-admin');
const momentTz = require('moment-timezone');
const currencyJs = require('currency.js');
const googleMapsClient = require('@google/maps').createClient({
  key: env.mapsApiKey,
  Promise: Promise,
});

const growthFileMsIntegration = async change => {
  if (!change.after.data()) {
    return null;
  }

  const {
    after: { id: activityId },
  } = change;
  const activityData = change.after.data();

  switch (activityData.template) {
    case 'office':
      activityData.officeId = activityId;
      activityData.template = 'office';
      break;
    case 'recipient':
      activityData.template = 'recipient';
      activityData.recipientId = activityId;
      break;
    default:
      return null;
  }

  return growthfileMsRequester({
    method: msRequestTypes.ACTIVITY,
    payload: activityData,
    resourcePath: msEndpoints.ACTIVITY,
  });
};

// updates the subscription activities on employee activity update
const handleSupervisorUpdate = async locals => {
  const { change } = locals;
  const activityEmployeeDataNew = change.after.data();

  // check for updated supervisors, delete old and new to all the employee's subscription activities
  const supervisors = ['First Supervisor'];
  const toDelete = [];
  const toAdd = [];
  const activityEmployeeOldAttachment = change.before.data().attachment;
  const activityEmployeeNewAttachment = activityEmployeeDataNew.attachment;

  supervisors.forEach(supervisor => {
    const oldNumber = activityEmployeeOldAttachment[supervisor].value;
    const newNumber = activityEmployeeNewAttachment[supervisor].value;
    if (oldNumber === newNumber) return;
    if (oldNumber !== '' && newNumber === '') {
      toDelete.push(oldNumber);
    }

    if (oldNumber === '' && newNumber !== '') {
      toAdd.push(newNumber);
    }

    if (oldNumber !== '' && newNumber !== '') {
      toDelete.push(oldNumber);
      toAdd.push(newNumber);
    }
  });

  // nothing to update, exit
  if (toDelete.length === 0 && toAdd.length === 0) return;

  // initiate a batch
  const batch = db.batch();

  // get all CONFIRMED subscriptions
  const { docs: subscriptionDocs } = await rootCollections.activities
    .where('officeId', '==', activityEmployeeDataNew.officeId)
    .where('template', '==', 'subscription')
    .where('status', '==', 'CONFIRMED')
    .where(
      'attachment.Phone Number.value',
      '==',
      activityEmployeeDataNew.attachment['Phone Number'].value,
    )
    // dont need to update check in subscription only
    // .where('attachment.Template.value', '==', 'check-in')
    .get();

  // filter out numbers uniquely
  const oldNumbers = [...new Set(toDelete)].filter(Boolean);
  const newNumbers = [...new Set(toAdd)].filter(Boolean);

  // update timestamp to force migration of this to profiles collection
  const updateTimeStamp = (batch, subscriptionDoc) => {
    batch.set(subscriptionDoc.ref, { timestamp: Date.now() }, { merge: true });
  };

  subscriptionDocs.forEach(subscriptionDoc => {
    // to prevent double timestamp update for scalability
    let isTimestampUpdated = false;
    oldNumbers.forEach(oldNumber => {
      isTimestampUpdated = true;
      updateTimeStamp(batch, subscriptionDoc);
      batch.delete(
        subscriptionDoc.ref
          .collection(subcollectionNames.ASSIGNEES)
          .doc(oldNumber),
      );
    });
    newNumbers.forEach(newNumber => {
      if (!isTimestampUpdated) updateTimeStamp(batch, subscriptionDoc);
      batch.set(
        subscriptionDoc.ref
          .collection(subcollectionNames.ASSIGNEES)
          .doc(newNumber),
        { addToInclude: true },
        { merge: true },
      );
    });
  });

  await batch.commit();
};

const getActivityObjectWithMetadata = doc =>
  Object.assign({}, doc.data(), {
    addendumDocRef: null,
    createTime: doc.createTime.toMillis(),
    id: doc.id,
    updateTime: doc.updateTime.toMillis(),
  });

const getProfileActivityObject = ({
  activityDoc,
  assigneesMap,
  assigneePhoneNumbersArray,
  customerObject,
}) => {
  const { id: activityId } = activityDoc;
  const dates = [];
  const schedules = activityDoc.get('schedule');
  if (schedules && Array.isArray(schedules) && schedules.length > 0) {
    schedules.forEach(({ startTime, endTime }) => {
      dates.push(
        ...enumerateDaysBetweenDates(startTime, endTime, dateFormats.DATE),
      );
    });
  }
  const intermediate = Object.assign({}, activityDoc.data(), {
    activityId,
    addendumDocRef: null,
    timestamp: Date.now(),
    dates: dates,
    assignees: assigneePhoneNumbersArray.map(phoneNumber => {
      const { displayName = '', photoURL = '' } =
        assigneesMap.get(phoneNumber) || {};

      return { phoneNumber, displayName, photoURL };
    }),
  });

  if (customerObject) {
    intermediate.customerObject = customerObject;
  }

  return intermediate;
};

const getUserRole = async ({ addendumDoc }) => {
  const { user: phoneNumber, roleDoc, activityData } = addendumDoc.data();
  const { officeId } = activityData;

  if (roleDoc) {
    return roleDoc;
  }

  /**
   * Destructuring the first element of the array
   * where the template is not `admin` OR `subscription`.
   */
  const [roleActivity] = (
    await rootCollections.activities
      .where('officeId', '==', officeId)
      .where('status', '==', 'CONFIRMED')
      .where('attachment.Phone Number.value', '==', phoneNumber)
      .get()
  ).docs.filter(doc => {
    const { template } = doc.data();

    // User's role activity
    // is employee currently.
    return template !== 'admin' && template !== 'subscription';
  });

  const batch = db.batch();
  const role = roleActivity
    ? getActivityObjectWithMetadata(roleActivity)
    : null;

  // Only the check-in subscription has `roleDoc` value by default.
  if (role) {
    batch.set(addendumDoc.ref, { roleDoc: role }, { merge: true });

    const {
      docs: [checkInSubscriptionDoc],
    } = await rootCollections.profiles
      .doc(phoneNumber)
      .collection(subcollectionNames.SUBSCRIPTIONS)
      .where('officeId', '==', officeId)
      .where('template', '==', 'check-in')
      .where('status', '==', 'CONFIRMED')
      .limit(1)
      .get();

    if (checkInSubscriptionDoc) {
      batch.set(checkInSubscriptionDoc.ref, { roleDoc: role }, { merge: true });
    }
  }

  await batch.commit();

  return role;
};

const getRoleReportData = (roleDocData, phoneNumber) => {
  if (!roleDocData) {
    return {
      phoneNumber,
      id: '',
      activationDate: '',
      employeeName: '',
      employeeCode: '',
      baseLocation: '',
      region: '',
      department: '',
      minimumDailyActivityCount: '',
      minimumWorkingHours: '',
      locationValidationCheck: '',
    };
  }

  return {
    phoneNumber,
    id: roleDocData.id,
    locationValidationCheck:
      roleDocData.attachment['Location Validation Check'].value,
    activationDate: roleDocData.createTime,
    employeeName: roleDocData.attachment.Name.value,
    employeeCode: roleDocData.attachment['Employee Code'].value,
    baseLocation: roleDocData.attachment['Base Location'].value,
    region: roleDocData.attachment.Region.value,
    department: roleDocData.attachment.Department.value,
    minimumDailyActivityCount:
      roleDocData.attachment['Minimum Daily Activity Count'].value,
    minimumWorkingHours: roleDocData.attachment['Minimum Working Hours'].value,
  };
};

const getActivityReportName = async ({ report, template }) => {
  if (report) {
    return report;
  }

  const {
    docs: [templateDoc],
  } = await rootCollections.activityTemplates
    .where('name', '==', template)
    .limit(1)
    .get();

  // Some old activities are present with templates
  // which have been deleted. Eg. expense-type.
  // Handling that here.
  const result = templateDoc ? templateDoc.get('report') : null;

  /**
   * Not all templates will have this field.
   */
  return result || null;
};

const getAttendanceHoursWorked = ({ attendanceData, date }) => {
  const checkInsArray = attendanceData.attendance[date].addendum || [];
  const { length: numberOfCheckIns } = checkInsArray;

  if (checkInsArray.length === 0) {
    return 0;
  }

  const [first] = attendanceData.attendance[date].addendum;
  const last = attendanceData.attendance[date].addendum[numberOfCheckIns - 1];

  return momentTz(momentTz(last)).diff(momentTz(first), 'hours', true);
};

const getLocationUrl = plusCode => `https://plus.codes/${plusCode}`;

const getLocalityCityState = mapsApiResult => {
  let locality = '';
  let city = '';
  let state = '';

  if (mapsApiResult.json.results.length === 0) {
    return {
      locality,
      city,
      state,
    };
  }

  const {
    json: {
      results: [{ address_components: components }],
    },
  } = mapsApiResult;

  components.forEach(component => {
    if (component.types.includes('locality')) {
      locality = component.long_name;
    }

    if (component.types.includes('administrative_area_level_2')) {
      city = component.long_name;
    }

    if (component.types.includes('administrative_area_level_1')) {
      state = component.long_name;
    }
  });

  return { locality, city, state };
};

const getPlaceInformation = (mapsApiResult, geopoint) => {
  const value = toMapsUrl(geopoint);

  if (!mapsApiResult) {
    return { url: value, identifier: value };
  }

  const {
    json: {
      results: [firstResult],
      plus_code: { global_code: globalCode },
    },
  } = mapsApiResult;

  if (!firstResult) {
    return { url: value, identifier: value };
  }

  return {
    identifier: firstResult.formatted_address,
    url: getLocationUrl(globalCode),
  };
};

const getUpdatedVenueDescriptors = (newVenue, oldVenue) => {
  const updatedFields = [];

  oldVenue.forEach((venue, index) => {
    const {
      venueDescriptor,
      location: oldLocation,
      address: oldAddress,
      geopoint: oldGeopoint,
    } = venue;

    const oldLongitude = oldGeopoint._longitude;
    const oldLatitude = oldGeopoint._latitude;
    const newLocation = newVenue[index].location;
    const newAddress = newVenue[index].address;
    const newGeopoint = newVenue[index].geopoint;
    const newLatitude = newGeopoint.latitude;
    const newLongitude = newGeopoint.longitude;

    if (
      oldLocation === newLocation &&
      oldAddress === newAddress &&
      oldLatitude === newLatitude &&
      oldLongitude === newLongitude
    ) {
      return;
    }

    updatedFields.push(venueDescriptor);
  });

  return updatedFields;
};

const getLocationObject = async ({ name, officeId }) => {
  if (!name) {
    return null;
  }

  const {
    docs: [locationDoc],
  } = await rootCollections.activities
    .where('attachment.Name.value', '==', name)
    .where('officeId', '==', officeId)
    .where('status', '==', 'CONFIRMED')
    .limit(1)
    .get();

  if (!locationDoc) {
    return null;
  }

  const { attachment } = locationDoc.data();
  const [venue] = locationDoc.get('venue');
  const { location, address, geopoint } = venue;

  const object = {
    address,
    location,
    latitude: geopoint.latitude || geopoint._latitude,
    longitude: geopoint.longitude || geopoint._longitude,
  };

  Object.keys(attachment).forEach(field => {
    object[field] = attachment[field].value;
  });

  return object;
};

const createAdmin = async (locals, adminContact) => {
  if (!adminContact || !locals.addendumDoc) {
    return;
  }

  const { officeId } = locals.change.after.data();

  const batch = db.batch();
  const activityRef = rootCollections.activities.doc();
  const addendumDocRef = rootCollections.offices
    .doc(officeId)
    .collection(subcollectionNames.ADDENDUM)
    .doc();

  const [
    {
      docs: [adminTemplateDoc],
    },
    adminQuery,
  ] = await Promise.all([
    rootCollections.activityTemplates
      .where('name', '==', 'admin')
      .limit(1)
      .get(),
    rootCollections.activities
      .where('attachment.Phone Number.value', '==', adminContact)
      .where('office', '==', officeId)
      .where('status', '==', 'CONFIRMED')
      .limit(1)
      .get(),
  ]);

  /** Is already an admin */
  if (!adminQuery.empty) {
    return;
  }

  const activityData = {
    officeId,
    addendumDocRef,
    office: locals.change.after.get('office'),
    timezone: locals.change.after.get('timezone'),
    timestamp: locals.addendumDocData.timestamp,
    schedule: [],
    venue: [],
    attachment: {
      'Phone Number': {
        value: adminContact,
        type: 'phoneNumber',
      },
    },
    canEditRule: adminTemplateDoc.get('canEditRule'),
    creator: locals.change.after.get('creator'),
    hidden: adminTemplateDoc.get('hidden'),
    status: adminTemplateDoc.get('statusOnCreate'),
    template: 'admin',
    activityName: `ADMIN: ${adminContact}`,
    createTimestamp: Date.now(),
  };

  const addendumDocData = {
    activityData,
    timestamp: Date.now(),
    user: locals.change.after.get('creator.phoneNumber'),
    userDisplayName: locals.change.after.get('creator.displayName'),
    action: httpsActions.create,
    template: 'admin',
    location: locals.addendumDoc.get('location'),
    userDeviceTimestamp: locals.addendumDoc.get('userDeviceTimestamp'),
    activityId: activityRef.id,
    activityName: `ADMIN: ${adminContact}`,
    isSupportRequest: locals.addendumDoc.get('isSupportRequest') || false,
    isAdminRequest: locals.addendumDoc.get('isAdminRequest') || false,
    geopointAccuracy: null,
    provider: null,
    isAutoGenerated: true,
  };

  batch.set(activityRef, activityData);
  batch.set(addendumDocRef, addendumDocData);

  locals.assigneePhoneNumbersArray.forEach(phoneNumber => {
    batch.set(
      activityRef.collection(subcollectionNames.ASSIGNEES).doc(phoneNumber),
      { addToInclude: false },
    );
  });

  return batch.commit();
};

const handleXTypeActivities = async locals => {
  const batch = db.batch();
  const template = locals.change.after.get('attachment.Template.value');
  const subscriber = locals.change.after.get('attachment.Phone Number.value');
  const { officeId } = locals.change.after.data();
  const typeActivities = await rootCollections.activities
    .where('officeId', '==', officeId)
    .where('template', '==', `${template}-type`)
    .where('status', '==', 'CONFIRMED')
    .get();

  // if subscription is created/updated
  // fetch all x-type activities from
  // Offices/(officeId)/Activities
  // Put those activities in the subscriber path
  // Profiles/(subscriber)/Activities/{x-type activityId}/
  typeActivities.forEach(activity => {
    batch.set(
      rootCollections.profiles
        .doc(subscriber)
        .collection(subcollectionNames.ACTIVITIES)
        .doc(activity.id),
      Object.assign({}, activity.data(), { addendumDocRef: null }),
      { merge: true },
    );
  });

  return batch.commit();
};

const handleCanEditRule = async (locals, templateDoc) => {
  if (templateDoc.get('canEditRule') !== 'ADMIN') {
    return;
  }

  const { office, officeId, status, attachment } = locals.change.after.data();
  const { value: subscriberPhoneNumber } = attachment['Phone Number'];

  if (status === 'CANCELLED') {
    const userSubscriptions = await rootCollections.profiles
      .doc(subscriberPhoneNumber)
      .collection(subcollectionNames.SUBSCRIPTIONS)
      .where('canEditRule', '==', 'ADMIN')
      .where('status', '==', 'CONFIRMED')
      .where('office', '==', office)
      .get();

    if (!userSubscriptions.empty) {
      return;
    }

    // cancel admin activity for this user
    const adminActivityQueryResult = await rootCollections.activities
      .where('status', '==', 'CONFIRMED')
      .where('template', '==', 'admin')
      .where('attachment.Phone Number.value', '==', subscriberPhoneNumber)
      .where('officeId', '==', officeId)
      .limit(1)
      .get();

    if (adminActivityQueryResult.empty) {
      return;
    }

    return adminActivityQueryResult.docs[0].ref.set(
      { status: 'CANCELLED', addendumDocRef: null },
      { merge: true },
    );
  }

  return createAdmin(locals, subscriberPhoneNumber);
};

const handleSubscription = async locals => {
  const batch = db.batch();
  const { id: activityId } = locals.change.after;
  const subscribedTemplate = locals.change.after.get(
    'attachment.Template.value',
  );
  const newSubscriber = locals.change.after.get(
    'attachment.Phone Number.value',
  );
  const oldSubscriber = locals.change.before.get(
    'attachment.Phone Number.value',
  );
  const subscriptionDocRef = rootCollections.profiles
    .doc(newSubscriber)
    .collection(subcollectionNames.SUBSCRIPTIONS)
    .doc(activityId);

  const promises = [
    rootCollections.activityTemplates
      .where('name', '==', subscribedTemplate)
      .limit(1)
      .get(),
  ];

  if (subscribedTemplate === 'check-in') {
    promises.push(
      rootCollections.activities
        .where('officeId', '==', locals.change.after.get('officeId'))
        .where('attachment.Phone Number.value', '==', newSubscriber)
        .where('status', '==', 'CONFIRMED')
        .get(),
    );
  }

  const [
    {
      docs: [templateDoc],
    },
    subscriberRoleDocs,
  ] = await Promise.all(promises);

  const include = locals.assigneePhoneNumbersArray.filter(phoneNumber => {
    return (
      newSubscriber !== phoneNumber &&
      locals.assigneesMap.get(phoneNumber).addToInclude
    );
  });

  const subscriptionDocData = {
    template: subscribedTemplate,
    timestamp: Date.now(),
    include: Array.from(new Set(include)),
    office: locals.change.after.get('office'),
    status: locals.change.after.get('status'),
  };

  if (subscriberRoleDocs && !subscriberRoleDocs.empty) {
    const [roleDoc] = subscriberRoleDocs.docs.filter(doc => {
      const { template } = doc.data();

      return template !== 'admin' && template !== 'subscription';
    });

    if (roleDoc) {
      subscriptionDocData.roleDoc = getActivityObjectWithMetadata(roleDoc);
    }
  }
  const subscriptionData = locals.change.after.data();
  //batch.set(subscriptionDocRef, subscriptionDocData, { merge: true });
  batch.set(
    subscriptionDocRef,
    Object.assign({}, subscriptionData, {
      include: Array.from(new Set(include)),
    }),
    { merge: true },
  );

  const newSubscriberAuth = await getAuth(newSubscriber);

  if (newSubscriberAuth.uid) {
    batch.set(
      rootCollections.updates
        .doc(newSubscriberAuth.uid)
        .collection(subcollectionNames.ADDENDUM)
        .doc(),
      Object.assign({}, subscriptionDocData, {
        _type: addendumTypes.SUBSCRIPTION,
        activityId: locals.change.after.id,
        attachment: templateDoc.get('attachment'),
        venue: templateDoc.get('venue'),
        schedule: templateDoc.get('schedule'),
        template: subscribedTemplate,
        report: templateDoc.get('report') || '',
      }),
      { merge: true },
    );
  }

  /**
   * Delete subscription doc from old profile
   * if the phone number has been changed in the
   * subscription activity.
   */
  const subscriberChanged =
    locals.change.before.data() && oldSubscriber !== newSubscriber;

  /**
   * Subscriber changed, so, deleting old doc in old `Updates`
   * This doc might not exist.
   **/
  if (newSubscriberAuth.uid && subscriberChanged) {
    batch.delete(
      rootCollections.updates
        .doc(newSubscriberAuth.uid)
        .collection(subcollectionNames.ADDENDUM)
        .doc(locals.change.after.id),
    );
  }

  if (subscriberChanged) {
    batch.delete(
      rootCollections.profiles
        .doc(oldSubscriber)
        .collection(subcollectionNames.SUBSCRIPTIONS)
        .doc(locals.change.after.id),
    );
  }

  await batch.commit();
  await handleCanEditRule(locals, templateDoc);

  // Leave-type, customer type etc.
  return handleXTypeActivities(locals);
};

const removeFromOfficeActivities = async locals => {
  const { status, office } = locals.change.after.data();

  /** Only remove when the status is `CANCELLED` */
  if (status !== 'CANCELLED') {
    return;
  }

  let oldStatus;

  if (locals.change.before.data()) {
    oldStatus = locals.change.before.get('status');
  }

  if (oldStatus && oldStatus === 'CANCELLED' && status === 'CANCELLED') {
    return;
  }

  const { value: activityPhoneNumber } = locals.change.after.get(
    'attachment.Phone Number',
  );

  const runQuery = (query, resolve, reject) => {
    return query
      .get()
      .then(docs => {
        if (docs.empty) {
          return 0;
        }

        const batch = db.batch();

        docs.forEach(doc => {
          const { template, status: activityStatus } = doc.data();

          /**
           * Not touching the same activity which causes this flow
           * to run. Allowing that will send the activityOnWrite
           * to an infinite spiral (probably).
           */
          if (doc.id === locals.change.after.id) {
            return;
          }

          // No point of recancelling the already cancelled activities.
          if (activityStatus === 'CANCELLED') {
            return;
          }

          const {
            value: phoneNumberInAttachment,
            // Using empty object fallback because `Phone Number` might
            // not exist in the object.
          } = doc.get('attachment.Phone Number') || {};

          // Cancelling admin to remove their custom claims.
          // Cancelling subscription to stop them from
          // creating new activities with that subscription
          if (
            new Set(['admin', 'subscription']).has(template) &&
            activityPhoneNumber === phoneNumberInAttachment
          ) {
            batch.set(
              rootCollections.activities.doc(doc.id),
              { status: 'CANCELLED', addendumDocRef: null },
              { merge: true },
            );

            return;
          }

          batch.set(
            rootCollections.activities.doc(doc.id),
            { addendumDocRef: null, timestamp: Date.now() },
            { merge: true },
          );

          batch.delete(
            rootCollections.activities
              .doc(doc.id)
              .collection(subcollectionNames.ASSIGNEES)
              .doc(activityPhoneNumber),
          );
        });

        /* eslint-disable */
        return batch.commit().then(() => docs.docs[docs.size - 1]);
        /* eslint-enable */
      })
      .then(lastDoc => {
        if (!lastDoc) {
          return resolve();
        }

        return process.nextTick(() => {
          const newQuery = query
            // Using greater than sign because we need
            // to start after the last activity which was
            // processed by this code otherwise some activities
            // might be updated more than once.
            .where(admin.firestore.FieldPath.documentId(), '>', lastDoc.id);

          return runQuery(newQuery, resolve, reject);
        });
      })
      .catch(new Error(reject));
  };

  const query = rootCollections.profiles
    .doc(activityPhoneNumber)
    .collection(subcollectionNames.ACTIVITIES)
    .where('office', '==', office)
    .orderBy('__name__')
    .limit(250);

  return new Promise((resolve, reject) =>
    runQuery(query, resolve, reject),
  ).catch(console.error);
};

const createDefaultSubscriptionsForUser = locals => {
  const { value: activityPhoneNumber } = locals.change.after.get(
    'attachment.Phone Number',
  );
  const { status } = locals.change.after.data();

  /**
   * Activity is cancelled, so creating subscription
   * is useless.
   */
  if (status === 'CANCELLED') {
    return;
  }

  return Promise.all([
    createAutoSubscription(locals, 'check-in', activityPhoneNumber),
    createAutoSubscription(locals, 'leave', activityPhoneNumber),
    createAutoSubscription(
      locals,
      'attendance regularization',
      activityPhoneNumber,
    ),
  ]);
};

const handleAttendanceDocs = async locals => {
  const batch = db.batch();
  const {
    officeId,
    status,
    timezone = 'Asia/Kolkata',
  } = locals.change.after.data();
  const { after: activityDoc } = locals.change;
  const { value: phoneNumber } = activityDoc.get('attachment.Phone Number');
  const momentNow = momentTz().tz(timezone);
  const month = momentNow.month();
  const year = momentNow.year();

  const {
    docs: [attendanceDoc],
  } = await rootCollections.offices
    .doc(officeId)
    .collection(subcollectionNames.ATTENDANCES)
    .where('month', '==', month)
    .where('year', '==', year)
    .where('phoneNumber', '==', phoneNumber)
    .limit(1)
    .get();

  /**
   * If employee cancellation date is not of current month
   * skip further code.
   */
  if (status === 'CANCELLED') {
    return;
  }

  const attendanceUpdate = attendanceDoc ? attendanceDoc.data() : {};
  attendanceUpdate.attendance = attendanceUpdate.attendance || {};

  const ref = attendanceDoc
    ? attendanceDoc.ref
    : rootCollections.offices
        .doc(officeId)
        .collection(subcollectionNames.ATTENDANCES)
        .doc();

  const { uid } = await getAuth(phoneNumber);

  batch.set(
    ref,
    Object.assign({}, attendanceUpdate, {
      month,
      year,
      phoneNumber,
      uid: uid || null,
      employeeName: activityDoc.get('attachment.Name.value') || '',
      employeeCode: activityDoc.get('attachment.Employee Code.value') || '',
      baseLocation: activityDoc.get('attachment.Base Location.value') || '',
      region: activityDoc.get('attachment.Region.value') || '',
      department: activityDoc.get('attachment.Department.value') || '',
      roleDoc: locals.roleObject || null,
    }),
    { merge: true },
  );

  return batch.commit();
};

const updateCheckInSubscriptionRoleField = async locals => {
  // This activity is most probably `employee`
  const batch = db.batch();
  const { after: activityDoc } = locals.change;
  const { officeId, attachment } = activityDoc.data();
  const { value: phoneNumber } = attachment['Phone Number'];

  const {
    docs: [checkInSubscriptionActivity],
  } = await rootCollections.activities
    .where('officeId', '==', officeId)
    .where('template', '==', 'subscription')
    .where('attachment.Phone Number.value', '==', phoneNumber)
    .where('attachment.Template.value', '==', 'check-in')
    .where('status', '==', 'CONFIRMED')
    .limit(1)
    .get();

  if (!checkInSubscriptionActivity) {
    return;
  }

  const { id: checkInId } = checkInSubscriptionActivity;

  // Manage doc below profile. Keeps it in sync with role activity.
  batch.set(
    rootCollections.profiles
      .doc(phoneNumber)
      .collection(subcollectionNames.SUBSCRIPTIONS)
      .doc(checkInId),
    { roleDoc: getActivityObjectWithMetadata(activityDoc) },
    { merge: true },
  );

  return batch.commit();
};

// handlRole
const handleConfig = async locals => {
  const batch = db.batch();
  const profileData = {};
  const { office, officeId, template } = locals.change.after.data();
  const { value: newActivityPhoneNumber } =
    locals.change.after.get('attachment.Phone Number') || {};
  const [venue] = locals.change.after.get('venue');
  const employeeOf = {
    [office]: officeId,
  };

  /**
   * check-in, recipient, and office don't have attachment.Phone Number
   * .value field
   */
  if (new Set(['admin', 'subscription']).has(template) || venue) {
    return;
  }

  if (!newActivityPhoneNumber) {
    return;
  }

  const hasBeenCancelled =
    locals.change.before.data() &&
    locals.change.before.get('status') !== 'CANCELLED' &&
    locals.change.after.get('status') === 'CANCELLED';
  const hasBeenCreated =
    locals.addendumDoc &&
    locals.addendumDoc.get('action') === httpsActions.create;

  // Change of status from `CONFIRMED` to `CANCELLED`
  if (hasBeenCancelled) {
    employeeOf[office] = admin.firestore.FieldValue.delete();
  }

  if (hasBeenCreated) {
    /**
     * New employee needs to have the locations map in their app.
     */
    profileData.lastLocationMapUpdateTimestamp = Date.now();
    profileData.employeeOf = employeeOf;
  }

  batch.set(rootCollections.profiles.doc(newActivityPhoneNumber), profileData, {
    merge: true,
  });

  await batch.commit();

  if (hasBeenCancelled) {
    await removeFromOfficeActivities(locals);
  }

  await createDefaultSubscriptionsForUser(locals);
  await handleAttendanceDocs(locals);

  return updateCheckInSubscriptionRoleField(locals);
};

const getUsersWithCheckInSubscription = async officeId => {
  const checkInSubscriptions = await rootCollections.activities
    .where('officeId', '==', officeId)
    .where('template', '==', 'subscription')
    .where('attachment.Template.value', '==', 'check-in')
    .where('status', '==', 'CONFIRMED')
    .get();

  return checkInSubscriptions.docs.map(doc =>
    doc.get('attachment.Phone Number.value'),
  );
};

// const mapActivityToUserUpdates = async (mappingActivity, prevQuery) => {
//   const { id: mappedActivityId } = mappingActivity;
//   const { officeId, venue, template, status } = mappingActivity.data();

//   // check-in template has venue, but its not a location activity.
//   if (template === 'check-in' || !venue[0] || status === 'CANCELLED') {
//     return;
//   }

//   const query =
//     prevQuery ||
//     rootCollections.activities
//       .where('officeId', '==', officeId)
//       .where('template', '==', 'subscription')
//       .where('attachment.Template.value', '==', 'check-in')
//       .where('status', '==', 'CONFIRMED')
//       .orderBy('__name__')
//       .limit(200);

//   const { docs, size, empty } = query.get();

//   if (empty) {
//     return;
//   }

//   const batch = db.batch();
//   const updateDocPromises = docs.map(doc => {
//     const { attachment } = doc.data();
//     const {
//       'Phone Number': { value: phoneNumber },
//     } = attachment;

//     return rootCollections.updates
//       .where('phoneNumber', '==', phoneNumber)
//       .limit(1)
//       .get();
//   });

//   const uidMap = new Map();

//   (await Promise.all(updateDocPromises)).forEach(({ docs }) => {
//     const [doc] = docs;

//     if (!doc) {
//       return;
//     }

//     const { uid } = doc;
//     const { phoneNumber } = doc.data();

//     uidMap.set(phoneNumber, uid);
//   });

//   docs.forEach(doc => {
//     const { attachment } = doc.data();
//     const {
//       'Phone Number': { value: phoneNumber },
//     } = attachment;

//     const { uid } = uidMap.get(phoneNumber);

//     batch.set(
//       rootCollections.updates
//         .doc(uid)
//         .collection(subcollectionNames.ADDENDUM)
//         .doc(),
//       Object.assign({}, location.data(), {
//         _type: addendumTypes.LOCATION,
//         activityId: mappedActivityId,
//         timestamp: Date.now(),
//       }),
//     );
//   });

//   const lastDoc = docs[size - 1];

//   return mapActivityToUserUpdates(mappingActivity, query.startAfter(lastDoc));
// };

const setLocationsReadEvent = async locals => {
  const { officeId, template } = locals.change.after.data();
  const [venue] = locals.change.after.get('venue');

  /**
   * Activity has no venue, so this activity is not
   * related to any location.
   */
  if (template === 'check-in' || !venue) {
    return;
  }

  let docsCounter = 0;
  let batchIndex = 0;
  const MAX_DOCS_ALLOWED_IN_A_BATCH = 500;
  const phoneNumbersArray = await getUsersWithCheckInSubscription(officeId);
  const numberOfBatches = Math.round(
    Math.ceil(phoneNumbersArray.length / MAX_DOCS_ALLOWED_IN_A_BATCH),
  );
  const batchArray = Array.from(Array(numberOfBatches)).map(() => db.batch());
  const updatesPromises = [];

  phoneNumbersArray.forEach(phoneNumber => {
    updatesPromises.push(
      rootCollections.updates
        .where('phoneNumber', '==', phoneNumber)
        .limit(1)
        .get(),
    );
  });

  const updateDocs = await Promise.all(updatesPromises);

  updateDocs.forEach(doc => {
    if (!doc.exists) {
      return;
    }

    docsCounter++;

    if (docsCounter > 499) {
      docsCounter = 0;
      batchIndex++;
    }

    batchArray[batchIndex].set(
      rootCollections.updates.doc(doc.id),
      { lastLocationMapUpdateTimestamp: Date.now() },
      { merge: true },
    );
  });

  const commitBatch = async batch => {
    return process.nextTick(() => batch.commit());
  };

  return batchArray.reduce(async (accumulatorPromise, currentBatch) => {
    await accumulatorPromise;

    return commitBatch(currentBatch);
  }, Promise.resolve());
};

const handleRecipient = async locals => {
  const batch = db.batch();
  const { id: activityId } = locals.change.after;
  const recipientsDocRef = rootCollections.recipients.doc(activityId);

  if (
    locals.addendumDoc &&
    locals.addendumDoc.get('action') === httpsActions.comment
  ) {
    return;
  }

  const { status, office, officeId, attachment } = locals.change.after.data();
  const { assigneePhoneNumbersArray: include } = locals;

  const {
    Name: { value: report },
    cc: { value: cc },
  } = attachment;
  const recipientObject = {
    status,
    include,
    office,
    officeId,
    report,
    cc,
  };

  batch.set(recipientsDocRef, recipientObject, { merge: true });

  if (status === 'CANCELLED') {
    batch.delete(recipientsDocRef);
  }

  await batch.commit();

  return connectionToExternal(
    Object.assign({}, recipientObject, { activityId }),
    'recipients',
  );
};

const createNewProfiles = async ({ newPhoneNumbersSet, smsContext }) => {
  const profileBatch = db.batch();
  const profilePromises = [];

  const promiseCreator = phoneNumber => {
    if (!phoneNumber) {return;}
    profilePromises.push(rootCollections.profiles.doc(phoneNumber).get());
  };

  newPhoneNumbersSet.forEach(promiseCreator);

  const snap = await Promise.all(profilePromises);
  const batchCreator = doc => {
    /** Profile already exists */
    if (doc.exists) {
      return;
    }

    // doc.id => phoneNumber
    profileBatch.set(doc.ref, { smsContext }, { merge: true });
  };

  snap.forEach(batchCreator);

  return profileBatch.commit();
};

const getCopyPath = ({ template, officeId, activityId }) => {
  if (template === 'office') {
    return rootCollections.offices.doc(activityId);
  }

  return rootCollections.offices
    .doc(officeId)
    .collection(subcollectionNames.ACTIVITIES)
    .doc(activityId);
};

/**
 * Checks if the action was a comment.
 * @param {string} action Can be one of the activity actions from HTTPS functions.
 * @returns {number} 0 || 1 depending on whether the action was a comment or anything else.
 */
const isComment = action => (action === httpsActions.comment ? 1 : 0);

const getUpdatedScheduleNames = (newSchedule, oldSchedule) => {
  const updatedFields = [];

  oldSchedule.forEach((item, index) => {
    const name = item.name;
    /** Request body ===> Update API request body. */
    const { startTime: newStartTime } = newSchedule[index];
    const { endTime: newEndTime } = newSchedule[index];
    const { startTime: oldStartTime } = item;
    const { endTime: oldEndTime } = item;

    if (newEndTime === oldEndTime && newStartTime === oldStartTime) {
      return;
    }

    updatedFields.push(name);
  });

  return updatedFields;
};

const getUpdatedAttachmentFieldNames = (newAttachment, oldAttachment) => {
  const updatedFields = [];

  Object.keys(newAttachment).forEach(field => {
    /** Comparing the `base64` photo string is expensive. Not doing it. */
    if (newAttachment[field].type === 'base64') {
      return;
    }

    const oldFieldValue = oldAttachment[field].value;
    const newFieldValue = newAttachment[field].value;
    const isUpdated = oldFieldValue !== newFieldValue;

    if (!isUpdated) {
      return;
    }

    updatedFields.push(field);
  });

  return updatedFields;
};

const getUpdatedFieldNames = options => {
  const { before: activityOld, after: activityNew } = options;
  const oldSchedule = activityOld.schedule;
  const oldVenue = activityOld.venue;
  const oldAttachment = activityOld.attachment;
  const newSchedule = activityNew.get('schedule');
  const newVenue = activityNew.get('venue');
  const newAttachment = activityNew.get('attachment');

  const allFields = [
    ...getUpdatedScheduleNames(newSchedule, oldSchedule),
    ...getUpdatedVenueDescriptors(newVenue, oldVenue),
    ...getUpdatedAttachmentFieldNames(newAttachment, oldAttachment),
  ];

  let commentString = '';

  if (allFields.length === 1) {
    return (commentString += `${allFields[0]}`);
  }

  allFields.forEach((field, index) => {
    if (index === allFields.length - 1) {
      commentString += `& ${field}`;

      return;
    }

    commentString += `${field}, `;
  });

  return commentString;
};

const getPronoun = (locals, recipient) => {
  const addendumCreator = locals.addendumDoc.get('user');
  const assigneesMap = locals.assigneesMap;

  if (addendumCreator === recipient) {
    return 'You';
  }

  if (
    assigneesMap.get(addendumCreator) &&
    assigneesMap.get(addendumCreator).displayName
  ) {
    return assigneesMap.get(addendumCreator).displayName;
  }

  if (
    !assigneesMap.get(addendumCreator) &&
    !locals.addendumCreatorInAssignees
  ) {
    return locals.addendumCreator.displayName;
  }

  /**
   * People are denoted with their phone numbers unless
   * the person creating the addendum is the same as the one
   * receiving it.
   */
  return addendumCreator;
};

const getCreateActionComment = (template, pronoun, locationFromVenue) => {
  const [templateNameFirstCharacter] = template;
  const article = vowels.has(templateNameFirstCharacter) ? 'an' : 'a';

  if (template === 'check-in' && locationFromVenue) {
    return `${pronoun} checked in from ${locationFromVenue}`;
  }

  return `${pronoun} created ${article} ${template}`;
};

const getChangeStatusComment = (status, activityName, pronoun) => {
  /** `PENDING` isn't grammatically correct with the comment here. */
  if (status === 'PENDING') {
    status = 'reversed';
  }

  return `${pronoun}` + ` ${status.toLowerCase()} ${activityName}`;
};

const getCommentString = (locals, recipient) => {
  const { action, share } = locals.addendumDoc.data();
  const { template } = locals.change.after.data();
  const pronoun = getPronoun(locals, recipient);

  if (action === httpsActions.create) {
    const locationFromVenue = (() => {
      if (template !== 'check-in') {
        return null;
      }

      if (
        locals.addendumDocData.activityData &&
        locals.addendumDocData.activityData.venue &&
        locals.addendumDocData.activityData.venue[0] &&
        locals.addendumDocData.activityData.venue[0].location
      ) {
        return locals.addendumDocData.activityData.venue[0].location;
      }

      if (locals.addendumDocData.venueQuery) {
        return locals.addendumDocData.venueQuery.location;
      }

      return locals.addendumDocData.identifier;
    })();

    return getCreateActionComment(template, pronoun, locationFromVenue);
  }

  if (action === httpsActions.changeStatus) {
    return getChangeStatusComment(
      locals.addendumDoc.get('status'),
      locals.addendumDoc.get('activityData.activityName'),
      pronoun,
    );
  }

  if (action === httpsActions.share) {
    let str = `${pronoun} added`;

    if (share.length === 1) {
      let name = locals.assigneesMap.get(share[0]).displayName || share[0];

      if (share[0] === recipient) name = 'you';

      return (str += ` ${name}`);
    }

    /** The `share` array will never have the `user` themselves */
    share.forEach((phoneNumber, index) => {
      let name =
        locals.assigneesMap.get(phoneNumber).displayName || phoneNumber;

      if (phoneNumber === recipient) {
        name = 'you';
      }

      if (share.length - 1 === index) {
        str += ` & ${name}`;

        return;
      }

      str += ` ${name}, `;
    });

    return str;
  }

  if (action === httpsActions.update) {
    const options = {
      before: locals.addendumDoc.get('activityOld'),
      after: locals.activityNew,
    };

    return `${pronoun} updated ${getUpdatedFieldNames(options)}`;
  }

  if (action === httpsActions.updatePhoneNumber) {
    const oldPhoneNumber = locals.addendumDoc.get('oldPhoneNumber');
    const newPhoneNumber = locals.addendumDoc.get('newPhoneNumber');

    // Employee changed their number themselves
    if (locals.addendumDoc.get('user') === oldPhoneNumber) {
      // <person name> changed their phone number
      // from < old phone number > to < new phone number >
      const userName = locals.addendumDoc.get(
        'activityData.attachment.Name.value',
      );

      return (
        `${userName} changed their phone number` +
        ` from ${oldPhoneNumber}` +
        ` to ${newPhoneNumber}`
      );
    }

    return (
      `Phone number` +
      ` '${oldPhoneNumber} was` +
      ` changed to ${newPhoneNumber}`
    );
  }

  /** Action is `comment` */
  return locals.addendumDoc.get('comment');
};

const getRegTokenMap = async assigneesMap => {
  const regTokenMap = new Map();
  const updateDocRefs = [];

  assigneesMap.forEach(({ uid }) => {
    if (!uid) {
      return;
    }

    updateDocRefs.push(rootCollections.updates.doc(uid));
  });

  /**
   * Need to check for empty array here because db.getAll throws
   * an error when passing 0 arguments.
   */
  if (updateDocRefs.length === 0) {
    return regTokenMap;
  }

  (await db.getAll(...updateDocRefs)).forEach(doc => {
    const { phoneNumber, registrationToken } = doc.data() || {};

    // registrationToken can be null or undefined or a
    // non-empty string.
    if (!registrationToken) {
      return;
    }

    regTokenMap.set(phoneNumber, registrationToken);
  });

  return regTokenMap;
};

const getNotificationObject = comment => ({
  data: {
    // Ask the client to send a request to the /read endpoint
    read: '1',
  },
  notification: {
    body: comment,
    tile: `Growthfile`,
  },
});

const getCommentObject = ({ addendumDoc, activityId, comment }) => ({
  comment,
  activityId,
  _type: addendumTypes.COMMENT,
  isComment: isComment(addendumDoc.get('action')),
  timestamp: addendumDoc.get('userDeviceTimestamp') || Date.now(),
  location: addendumDoc.get('location') || '',
  user: addendumDoc.get('user'),
});

const handleComments = async (addendumDoc, locals) => {
  if (!addendumDoc) {
    return;
  }

  const regTokenMap = await getRegTokenMap(locals.assigneesMap);
  const batch = db.batch();
  const notificationPromises = [];

  locals.assigneePhoneNumbersArray.forEach(phoneNumber => {
    const { uid } = locals.assigneesMap.get(phoneNumber);
    const registrationToken = regTokenMap.get(phoneNumber);

    if (!uid || !registrationToken) {
      return;
    }

    const comment = getCommentString(locals, phoneNumber);

    batch.set(
      rootCollections.updates
        .doc(uid)
        .collection(subcollectionNames.ADDENDUM)
        .doc(),
      getCommentObject({
        comment,
        addendumDoc,
        activityId: locals.change.after.id,
      }),
    );

    notificationPromises.push(
      admin
        .messaging()
        .sendToDevice(registrationToken, getNotificationObject(comment), {
          priority: 'high',
          timeToLive: 60,
        }),
    );
  });

  await batch.commit();

  return Promise.all(notificationPromises);
};

const createActivityStats = async addendumDoc => {
  if (!env.isProduction) {
    return;
  }

  const batch = db.batch();
  const getValue = (snap, field) => {
    if (snap.empty) {
      return 0;
    }

    return snap.docs[0].get(field) || 0;
  };

  const office = addendumDoc.get('activityData.office');
  const action = addendumDoc.get('action');
  const isSupportRequest = addendumDoc.get('isSupportRequest');
  const isAdminRequest = addendumDoc.get('isAdminRequest');
  const isAutoGenerated = addendumDoc.get('isAutoGenerated');
  const template = addendumDoc.get('activityData.template');
  const momentToday = momentTz().toObject();

  const [todayInitQuery, counterDocsQuery] = await Promise.all([
    rootCollections.inits
      .where('report', '==', reportNames.DAILY_STATUS_REPORT)
      .where('date', '==', momentToday.date)
      .where('month', '==', momentToday.months)
      .where('year', '==', momentToday.years)
      .limit(1)
      .get(),
    rootCollections.inits
      .where('report', '==', reportNames.COUNTER)
      .limit(1)
      .get(),
  ]);

  const initDocRef = snapShot => {
    return snapShot.empty ? rootCollections.inits.doc() : snapShot.docs[0].ref;
  };

  const initDoc = initDocRef(todayInitQuery);
  let totalActivities = counterDocsQuery.docs[0].get('totalActivities');
  let totalCreatedWithAdminApi = counterDocsQuery.docs[0].get(
    'totalCreatedWithAdminApi',
  );
  let totalCreatedWithClientApi = counterDocsQuery.docs[0].get(
    'totalCreatedWithClientApi',
  );
  let totalCreatedWithSupport = counterDocsQuery.docs[0].get(
    'totalCreatedWithSupport',
  );
  const supportMap = counterDocsQuery.docs[0].get('supportMap');
  const autoGeneratedMap = counterDocsQuery.docs[0].get('autoGeneratedMap');
  const totalByTemplateMap = counterDocsQuery.docs[0].get('totalByTemplateMap');
  const adminApiMap = counterDocsQuery.docs[0].get('adminApiMap');

  let activitiesAddedToday = getValue(todayInitQuery, 'activitiesAddedToday');
  let withAdminApi = getValue(todayInitQuery, 'withAdminApi');
  let autoGenerated = getValue(todayInitQuery, 'autoGenerated');
  let withSupport = getValue(todayInitQuery, 'withSupport');
  let createApi = getValue(todayInitQuery, 'createApi');
  let updateApi = getValue(todayInitQuery, 'updateApi');
  let changeStatusApi = getValue(todayInitQuery, 'changeStatusApi');
  let shareApi = getValue(todayInitQuery, 'shareApi');
  let commentApi = getValue(todayInitQuery, 'commentApi');

  const createCountByOffice = (() => {
    if (todayInitQuery.empty) {
      return {};
    }

    return todayInitQuery.docs[0].get('createCountByOffice') || {};
  })();

  if (action === httpsActions.create) {
    totalActivities++;
    activitiesAddedToday++;
    createApi++;

    if (!isSupportRequest && !isAdminRequest) {
      totalCreatedWithClientApi++;
    }

    if (totalByTemplateMap[template]) {
      totalByTemplateMap[template]++;
    } else {
      totalByTemplateMap[template] = 1;
    }

    if (createCountByOffice[office]) {
      createCountByOffice[office]++;
    } else {
      createCountByOffice[office] = 1;
    }
  }

  if (action === httpsActions.update) {
    updateApi++;
  }

  if (action === httpsActions.changeStatus) {
    changeStatusApi++;
  }

  if (action === httpsActions.share) {
    shareApi++;
  }

  if (action === httpsActions.comment) {
    commentApi++;
  }

  if (isSupportRequest) {
    withSupport++;
    totalCreatedWithSupport++;

    if (supportMap[template]) {
      supportMap[template]++;
    } else {
      supportMap[template] = 1;
    }
  }

  if (isAutoGenerated) {
    autoGenerated++;

    if (autoGeneratedMap[template]) {
      autoGeneratedMap[template]++;
    } else {
      autoGeneratedMap[template] = 1;
    }
  }

  if (isAdminRequest && !isSupportRequest) {
    // Support requests on admin resource does not count
    // towards this count.
    withAdminApi++;
    totalCreatedWithAdminApi++;

    if (adminApiMap[template]) {
      adminApiMap[template]++;
    } else {
      adminApiMap[template] = 1;
    }
  }

  const dataObject = todayInitQuery.empty ? {} : todayInitQuery.docs[0].data();

  dataObject.totalActivities = totalActivities;
  dataObject.activitiesAddedToday = activitiesAddedToday;
  dataObject.withAdminApi = withAdminApi;
  dataObject.autoGenerated = autoGenerated;
  dataObject.withSupport = withSupport;
  dataObject.createApi = createApi;
  dataObject.updateApi = updateApi;
  dataObject.changeStatusApi = changeStatusApi;
  dataObject.shareApi = shareApi;
  dataObject.commentApi = commentApi;
  dataObject.report = reportNames.DAILY_STATUS_REPORT;
  dataObject.date = new Date().getDate();
  dataObject.month = new Date().getMonth();
  dataObject.year = new Date().getFullYear();
  dataObject.createCountByOffice = createCountByOffice;

  if (!dataObject.templateUsageObject) {
    dataObject.templateUsageObject = {};
  }

  if (!dataObject.templateUsageObject[template]) {
    dataObject.templateUsageObject[template] = {};
  }

  if (!dataObject.templateUsageObject[template][action]) {
    dataObject.templateUsageObject[template][action] = 0;
  }

  dataObject.templateUsageObject[template][action] =
    dataObject.templateUsageObject[template][action] + 1;

  batch.set(initDoc, dataObject, { merge: true });

  // Counter always exists because it has been created manually
  // for storing counts of stuff...
  batch.set(
    counterDocsQuery.docs[0].ref,
    {
      totalActivities,
      adminApiMap,
      autoGeneratedMap,
      supportMap,
      totalByTemplateMap,
      totalCreatedWithAdminApi,
      totalCreatedWithClientApi,
      totalCreatedWithSupport,
    },
    { merge: true },
  );

  return batch.commit();
};

const getAccuracyTolerance = accuracy => {
  /**
   * Client sends geopointAccuracy in meteres.
   */
  if (accuracy && accuracy < 350) {
    return 0.5;
  }

  return 1;
};

const getVenueFromActivity = activityData => {
  const [venue] = activityData.venue;

  /**
   * Some activities might have empty venue object
   * based on template.
   */
  return venue ? venue : null;
};

const handlePopulatedVenue = ({ addendumDoc }) => {
  /** User's current location */
  const deviceLocation = {
    _latitude: addendumDoc.get('location')._latitude,
    _longitude: addendumDoc.get('location')._longitude,
  };

  const { activityData, geopointAccuracy } = addendumDoc.data();
  const activityVenue = getVenueFromActivity(activityData);
  const distanceTolerance = getAccuracyTolerance(geopointAccuracy);

  // venue is populated => calculate distabce btw actual and venue location.
  // haversineDistance(geopointOne, geopointTwo)
  return {
    distanceAccurate:
      haversineDistance(deviceLocation, activityVenue.geopoint) <
      distanceTolerance,
    venueQuery: activityVenue,
  };
};

const handleUnpopulatedVenue = async ({ addendumDoc }) => {
  // venue is not populated.
  // query db with adjusted geopoint.
  const {
    activityData,
    geopointAccuracy,
    location: currentGeopoint,
  } = addendumDoc.data();
  const { officeId } = activityData;
  const distanceTolerance = getAccuracyTolerance(geopointAccuracy);
  const adjGP = adjustedGeopoint(currentGeopoint);

  const {
    docs: [queriedActivity],
  } = await rootCollections.activities
    .where('officeId', '==', officeId)
    // Branch, and customer
    .where('adjustedGeopoints', '==', `${adjGP.latitude},${adjGP.longitude}`)
    .where('status', '==', 'CONFIRMED')
    .limit(1)
    .get();

  // { isAccurate: false, venueQuery: null };
  if (!queriedActivity) {
    const mapsApiResult = await googleMapsClient
      .reverseGeocode({
        latlng: getLatLngString(currentGeopoint),
      })
      .asPromise();

    const csl = getLocalityCityState(mapsApiResult);
    const ui = getPlaceInformation(mapsApiResult, currentGeopoint);

    return Object.assign({}, csl, ui, {
      distanceAccurate: false,
      venueQuery: null,
    });
  }

  const [activityVenue] = queriedActivity.get('venue');
  const { geopoint: activityGeopoint } = activityVenue;

  const distanceBetween = haversineDistance(currentGeopoint, activityGeopoint);

  return {
    distanceAccurate: distanceBetween < distanceTolerance,
    venueQuery: activityVenue,
  };
};

const handleNoVenue = async ({ currentGeopoint }) => {
  const mapsApiResult = await googleMapsClient
    .reverseGeocode({
      latlng: getLatLngString(currentGeopoint),
    })
    .asPromise();

  /** city, state, locality */
  const csl = getLocalityCityState(mapsApiResult);
  /** url, identifier */
  const ui = getPlaceInformation(mapsApiResult, currentGeopoint);

  return Object.assign({}, csl, ui, {
    isAccurate: false,
    venueQuery: null,
  });
};

const checkDistanceAccurate = async ({ addendumDoc }) => {
  //   if activity.venue is NOT populated => use adjustedGeopoint and query / Activities collection;
  //   if a doc is found:
  //   check haversine distance between queried activity geopoint and current geopoint
  //   if distance < 1km:
  //     distance accurate = true
  //   else:
  //     distance accurate = false

  const { activityData, location: currentGeopoint } = addendumDoc.data();
  const activityVenue = getVenueFromActivity(activityData);

  /** Activity with template that does have venue array of 0 length */
  if (!activityVenue) {
    return handleNoVenue({
      currentGeopoint,
    });
  }

  if (activityVenue.location) {
    return handlePopulatedVenue({
      addendumDoc,
    });
  }

  return handleUnpopulatedVenue({
    addendumDoc,
  });
};

const setGeopointAndTimestampInCheckInSubscription = async ({
  addendumDoc,
}) => {
  const {
    activityData: { office },
    user: phoneNumber,
    location: lastGeopoint,
    timestamp: lastTimestamp,
  } = addendumDoc.data();

  const [checkInSubscriptionDoc] = (
    await rootCollections.profiles
      .doc(phoneNumber)
      .collection(subcollectionNames.SUBSCRIPTIONS)
      .where('office', '==', office)
      .where('template', '==', 'check-in')
      .where('status', '==', 'CONFIRMED')
      .limit(1)
      .get()
  ).docs;

  if (!checkInSubscriptionDoc) {
    return;
  }

  return checkInSubscriptionDoc.ref.set(
    {
      lastGeopoint,
      lastTimestamp,
      lastAddendumRef: addendumDoc.ref,
    },
    { merge: true },
  );
};

const getDistanceTravelled = ({
  previousAddendumDoc,
  distanceMatrixApiResult,
}) => {
  if (!previousAddendumDoc) {
    return 0;
  }

  const [firstRow] = distanceMatrixApiResult.json.rows;
  const [firstElement] = firstRow.elements;
  const { distance: distanceData } = firstElement;

  return distanceData ? distanceData.value / 1000 : 0;
};

const handleAddendum = async locals => {
  const { addendumDoc } = locals;

  if (!addendumDoc) {
    return;
  }

  const {
    action,
    timestamp,
    location: currentGeopoint,
    user: phoneNumber,
  } = addendumDoc.data();

  const momentWithOffset = momentTz(timestamp).tz(
    addendumDoc.get('activityData.timezone') || 'Asia/Kolkata',
  );

  const isSkippableEvent =
    action === httpsActions.install ||
    action === httpsActions.signup ||
    action === httpsActions.branchView ||
    action === httpsActions.productView ||
    action === httpsActions.videoPlay;
  const date = momentWithOffset.date();
  const month = momentWithOffset.month();
  const year = momentWithOffset.year();

  if (isSkippableEvent) {
    return addendumDoc.ref.set({ date, month, year }, { merge: true });
  }

  /** Phone Number change addendum does not have geopoint */
  if (!currentGeopoint) {
    return;
  }

  const batch = db.batch();
  const promises = [];
  let previousGeopoint;

  // venue populated => fetch branch/customer activity with adjustedGeopoints {lat,lng}
  // if doc is found:
  // distanceAccurate => diff CURR and resulting activity < 1
  // if doc is not found:
  // distance accurate = false.

  //   if activity.venue is NOT populated => use adjustedGeopoint and query / Activities collection;
  //   if a doc is found:
  //   check haversine distance between queried activity geopoint and current geopoint
  //   if distance < 1km:
  //     distance accurate = true
  //   else:
  //     distance accurate = false

  const addendumQuery = await rootCollections.offices
    .doc(locals.change.after.get('officeId'))
    .collection(subcollectionNames.ADDENDUM)
    .where('user', '==', phoneNumber)
    .where('timestamp', '<', timestamp)
    .orderBy('timestamp', 'desc')
    .limit(2)
    .get();

  const previousAddendumDoc = (() => {
    if (addendumQuery.docs[0] && addendumQuery.docs[0].id !== addendumDoc.id) {
      return addendumQuery.docs[0];
    }

    return addendumQuery.docs[1];
  })();

  if (previousAddendumDoc) {
    /**
     * The field `location` could be undefined for install or
     * signup events in the previous addendum
     */
    previousGeopoint = previousAddendumDoc.get('location') || currentGeopoint;

    promises.push(
      googleMapsClient
        .distanceMatrix({
          /**
           * Ordering is important here. The `legal` distance
           * between A to B might not be the same as the legal
           * distance between B to A. So, do not mix the ordering.
           */
          // @ts-ignore
          origins: getLatLngString(previousGeopoint),
          // @ts-ignore
          destinations: getLatLngString(currentGeopoint),
          units: 'metric',
        })
        .asPromise(),
    );
  }

  const [distanceMatrixApiResult] = await Promise.all(promises);

  const daR = await checkDistanceAccurate({
    addendumDoc,
  });
  const updateObject = Object.assign({}, daR, {
    date,
    month,
    year,
    customerObject: locals.customerObject,
    // [{ displayName, phoneNumber, photoURL }]
    assigneesMap: [...locals.assigneesMap.values()],
    adjustedGeopoint: adjustedGeopoint(addendumDoc.get('location')),
    distanceTravelled: getDistanceTravelled({
      previousAddendumDoc,
      distanceMatrixApiResult,
    }),
  });

  // Required for comment creation since the addendumDoc.data() won't contain
  // the updates made during this function instance
  locals.addendumDocData = Object.assign({}, addendumDoc.data(), updateObject);
  // Used in reimburseKmAllowance function for getting the previous location
  locals.previousAddendumDoc = previousAddendumDoc;

  /**
   * Seperating this part out because handling even a single crash
   * with `addendumOnCreate` cloud function messes up whole data for the user
   * after the time of the crash.
   */
  batch.set(addendumDoc.ref, updateObject, { merge: true });

  await batch.commit();
  await setGeopointAndTimestampInCheckInSubscription({ addendumDoc });

  locals.roleObject = await getUserRole({ addendumDoc });

  return createActivityStats(addendumDoc);
};

const getMetaBaseQuery = ({ officeId, name, template }) => {
  const baseQuery = rootCollections.activities
    .where('officeId', '==', officeId)
    .where('template', '==', 'employee');

  if (template === 'branch') {
    return baseQuery.where('attachment.Base Location.value', '==', name);
  }

  if (template === 'region') {
    return baseQuery.where('attachment.Region.value', '==', name);
  }

  if (template === 'department') {
    return baseQuery.where('attachment.Department.value', '==', name);
  }

  return null;
};

/**
 * If name is updated
 * get all the activities with this name
 * and update the activities
 * If this instance has run because of activity being cancelled
 * during status-change, set all the activities using this value
 * in their type as '' (empty string).
 */
const handleMetaUpdate = async locals => {
  if (
    !locals.change.before.data() ||
    !locals.change.after.get('attachment.Name.value')
  ) {
    return;
  }

  const {
    template,
    officeId,
    status: newStatus,
    attachment: {
      Name: { value: newName },
    },
  } = locals.change.after.data();
  const {
    attachment: {
      Name: { value: oldName },
    },
  } = locals.change.before.data();

  /**
   * Name was not changed, so no need to proceed further
   */
  if (oldName && oldName === newName) {
    return;
  }

  const query = getMetaBaseQuery({
    officeId,
    template,
    name: oldName,
  });

  // Only proceed for branch, region and department
  if (!query) {
    return;
  }

  const value = (() => {
    if (newStatus === 'CANCELLED') {
      return '';
    }

    return newName;
  })();

  const field = (() => {
    if (template === 'branch') {
      return 'Base Location';
    }
    if (template === 'region') {
      return 'Region';
    }
    if (template === 'department') {
      return 'Department';
    }
  })();

  const docs = await query.get();
  const MAX_DOCS_ALLOWED_IN_A_BATCH = 500;
  const numberOfBatches = Math.round(
    Math.ceil(docs.size / MAX_DOCS_ALLOWED_IN_A_BATCH),
  );
  const batchArray = Array.from(Array(numberOfBatches)).map(() => db.batch());
  let batchIndex = 0;
  let docsCounter = 0;

  docs.forEach(doc => {
    if (docsCounter > 499) {
      docsCounter = 0;
      batchIndex++;
    }

    docsCounter++;

    batchArray[batchIndex].set(
      doc.ref,
      {
        addendumDocRef: null,
        attachment: {
          [field]: { value },
        },
      },
      { merge: true },
    );
  });

  return Promise.all(batchArray.map(batch => batch.commit()));
};

const getDutyCheckIns = ({ doc, phoneNumber }) => {
  const checkIns = doc.get('checkIns') || {};

  checkIns[phoneNumber] = checkIns[phoneNumber] || [];
  checkIns[phoneNumber].push(Date.now());

  return checkIns;
};

const handleScheduledActivities = async locals => {
  const batch = db.batch();
  const { officeId, timezone } = locals.change.after.data();
  const momentNow = momentTz().tz(timezone);
  const { phoneNumber, displayName } = locals.change.after.get('creator');

  // const [
  //   todaysActivities,
  //   yesterdayActivities,
  //   tomorrowActivities,
  // ] = await Promise.all([
  //   rootCollections
  //   .profiles
  //   .doc(phoneNumber)
  //   .collection(subcollectionNames.ACTIVITIES)
  //   .where('officeId', '==', officeId)
  //   .where('scheduleDates', 'array-contains', dateStringToday)
  //   .get(),
  //   rootCollections
  //   .profiles
  //   .doc(phoneNumber)
  //   .collection(subcollectionNames.ACTIVITIES)
  //   .where('officeId', '==', officeId)
  //   .where('scheduleDates', 'array-contains', momentNow.clone().subtract(1, 'day').format(dateFormats.DATE))
  //   .get(),
  //   rootCollections
  //   .profiles
  //   .doc(phoneNumber)
  //   .collection(subcollectionNames.ACTIVITIES)
  //   .where('officeId', '==', officeId)
  //   .where('scheduleDates', 'array-contains', momentNow.clone().add(1, 'day').format(dateFormats.DATE))
  //   .get(),
  // ]);

  // const scheduledActivities = [].concat(
  //   ...todaysActivities.docs,
  //   ...yesterdayActivities.docs,
  //   ...tomorrowActivities.docs
  // );

  const scheduledActivities = await rootCollections.profiles
    .doc(phoneNumber)
    .collection(subcollectionNames.ACTIVITIES)
    .where('officeId', '==', officeId)
    .where(
      'scheduleDates',
      'array-contains',
      momentNow.format(dateFormats.DATE),
    )
    .get();

  const activityIds = new Set();

  scheduledActivities.forEach(doc => {
    const { id: activityId } = doc;
    const location = doc.get('attachment.Location.value');

    if (!location) {
      return;
    }

    const hd = haversineDistance(
      {
        _latitude: doc.get('customerObject.latitude'),
        _longitude: doc.get('customerObject.longitude'),
      },
      locals.addendumDocData.location,
    );

    if (hd > 1) {
      return;
    }

    /**
     * Since we are using an array_contains query
     * to fetch the activities, it is certainly possible that
     * we might get duplicate activities in a single response.
     * We should skip updating a single activity multiple times
     * because creating a comment will be redundant.
     */
    if (activityIds.has(activityId)) {
      return;
    }

    activityIds.add(activityId);

    const addendumDocRef = rootCollections.offices
      .doc(officeId)
      .collection(subcollectionNames.ADDENDUM)
      .doc();

    batch.set(addendumDocRef, {
      date: momentNow.date(),
      month: momentNow.month(),
      year: momentNow.year(),
      user: phoneNumber,
      action: httpsActions.checkIn,
      location: locals.addendumDoc.get('location'),
      timestamp: Date.now(),
      userDeviceTimestamp: locals.addendumDoc.get('userDeviceTimestamp'),
      isSupportRequest: locals.addendumDoc.get('isSupportRequest'),
      geopointAccuracy: locals.addendumDoc.get('geopointAccuracy'),
      provider: locals.addendumDoc.get('provider'),
      userDisplayName: displayName,
      isAutoGenerated: true,
      comment:
        `${displayName || phoneNumber} checked ` +
        `in from ${doc.get('template')} Location: ${location}`,
      activityData: doc.data(),
      activityId: doc.ref.id,
    });

    batch.set(
      locals.change.after.ref,
      { addendumDocRef, timestamp: Date.now() },
      { merge: true },
    );

    batch.set(
      rootCollections.activities.doc(doc.id),
      {
        addendumDocRef,
        timestamp: Date.now(),
        checkIns: getDutyCheckIns({
          doc,
          phoneNumber,
        }),
      },
      { merge: true },
    );
  });

  // TODO: create additional comment
  // is not cancelled activity => leave
  // example comment => conflict with leave
  // only future leaves

  return batch.commit();
};

const copyTypeActivityToUserProfile = async (
  prevQuery,
  typeActivity,
  officeId,
  template,
) => {
  const query =
    prevQuery ||
    rootCollections.activities
      .where('officeId', '==', officeId)
      .where('template', '==', 'subscription')
      .where('attachment.Template.value', '==', template)
      .where('status', 'in', ['CONFIRMED', 'PENDING'])
      .orderBy('__name__')
      .limit(150);

  const docs = await query.get();

  if (docs.empty) {
    return;
  }

  const batch = db.batch();
  const uidMap = new Map();

  const authSnaps = await Promise.all(
    docs.docs.map(doc =>
      rootCollections.updates
        .where('phoneNumber', '==', doc.get('attachment.Phone Number.value'))
        .limit(1)
        .get(),
    ),
  );

  authSnaps.forEach(snap => {
    const {
      docs: [doc],
    } = snap;

    if (!doc) {
      return;
    }

    const { id: uid } = doc;
    const { phoneNumber } = doc.data();

    uidMap.set(phoneNumber, uid);
  });

  const baseObject = Object.assign({}, typeActivity.data(), {
    addendumDocRef: null,
    timestamp: Date.now(),
    activityId: typeActivity.id,
  });

  docs.forEach(activityDoc => {
    const phoneNumber = activityDoc.get('attachment.Phone Number.value');
    const uid = uidMap.get(phoneNumber);
    const { id: activityId } = activityDoc;

    batch.set(
      rootCollections.profiles
        .doc(phoneNumber)
        .collection(subcollectionNames.ACTIVITIES)
        .doc(activityId),
      baseObject,
    );

    if (!uid) {
      return;
    }

    batch.set(
      rootCollections.updates
        .doc(uid)
        .collection(subcollectionNames.ADDENDUM)
        .doc(),
      Object.assign({}, baseObject, {
        _type: addendumTypes.ACTIVITY,
      }),
    );
  });

  await batch.commit();
  const lastDoc = docs.docs[docs.size - 1];

  return copyTypeActivityToUserProfile(
    query.startAfter(lastDoc.id),
    typeActivity,
    officeId,
    template,
  );
};

const handleTypeActivity = async locals => {
  if (
    locals.addendumDoc &&
    (locals.addendumDoc.get('action') === httpsActions.comment ||
      locals.addendumDoc.get('action') === httpsActions.share)
  ) {
    return;
  }

  const { after: typeActivity } = locals.change;
  const { template, officeId } = locals.change.after.data();

  if (!template.endsWith('-type')) {
    return;
  }

  // eg => leave-type -> 'leave'
  const [parentTemplate] = template.split('-type');

  return copyTypeActivityToUserProfile(
    null,
    typeActivity,
    officeId,
    parentTemplate,
  );
};

const getReimbursementTimestamp = activityDoc => {
  const { template } = activityDoc.data();

  // For claim, if the schedule timestamp is present, that timestamp
  // will be the claim timestamp
  // otherwise, activity create time is the fallback.
  if (
    template === 'claim' &&
    activityDoc.get('schedule')[0] &&
    Number.isInteger(activityDoc.get('schedule')[0].startTime)
  ) {
    return activityDoc.get('schedule')[0].startTime;
  }

  return activityDoc.createTime.toDate().getTime();
};

const reimburseClaim = async locals => {
  const {
    office,
    officeId,
    status,
    timezone,
    template,
    creator: { phoneNumber },
  } = locals.change.after.data();
  const { id: activityId } = locals.change.after;
  const timestamp = getReimbursementTimestamp(locals.change.after);
  const momentNow = momentTz(timestamp).tz(timezone);
  const batch = db.batch();
  const date = momentNow.date();
  const month = momentNow.month();
  const year = momentNow.year();
  let uid = locals.addendumDocData.uid;

  if (!uid) {
    uid = (await getAuth(locals.addendumDocData.user)).uid;
  }

  const {
    docs: [reimsToday],
  } = await rootCollections.offices
    .doc(officeId)
    .collection(subcollectionNames.REIMBURSEMENTS)
    .where('date', '==', date)
    .where('month', '==', month)
    .where('year', '==', year)
    .where('phoneNumber', '==', phoneNumber)
    .where('claimId', '==', activityId)
    .limit(1)
    .get();

  const reimbursementData = Object.assign(
    {},
    getRoleReportData(locals.roleObject, phoneNumber),
    {
      status,
      date,
      month,
      year,
      office,
      officeId,
      uid,
      claimId: activityId,
      currency: 'INR',
      timestamp: Date.now(),
      reimbursementType: template,
      relevantActivityId: activityId,
      reimbursementName: locals.change.after.get('attachment.Claim Type.value'),
      photoURL: locals.change.after.get('attachment.Photo URL.value'),
      amount: currencyJs(
        locals.change.after.get('attachment.Amount.value'),
      ).toString(),
      claimType: locals.change.after.get('attachment.Claim Type.value'),
      activityDoc: locals.change.after.data(),
    },
  );

  if (locals.addendumDocData.action === httpsActions.changeStatus) {
    if (status === 'CANCELLED') {
      reimbursementData.cancelledBy = locals.addendumDocData.user;
      reimbursementData.cancellationTimestamp =
        locals.addendumDocData.timestamp;
    }

    if (status === 'CONFIRMED') {
      reimbursementData.confirmedBy = locals.addendumDocData.user;
      reimbursementData.confirmationTimestamp =
        locals.addendumDocData.timestamp;
    }
  }

  const reimbursementRef = reimsToday
    ? reimsToday.ref
    : rootCollections.offices
        .doc(officeId)
        .collection(subcollectionNames.REIMBURSEMENTS)
        .doc();

  batch.set(reimbursementRef, reimbursementData, { merge: true });

  // During phone number change, uid might not exist for the new number
  if (!uid) {
    return;
  }

  const {
    docs: [claimUpdatesDoc],
  } = await rootCollections.updates
    .doc(uid)
    .collection(subcollectionNames.ADDENDUM)
    .where('details.claimId', '==', activityId)
    .limit(1)
    .get();

  batch.set(
    claimUpdatesDoc
      ? claimUpdatesDoc.ref
      : rootCollections.updates
          .doc(uid)
          .collection(subcollectionNames.ADDENDUM)
          .doc(),
    {
      officeId,
      phoneNumber,
      date,
      month,
      year,
      office,
      activityId,
      timestamp: Date.now(),
      _type: addendumTypes.REIMBURSEMENT,
      amount: locals.change.after.get('attachment.Amount.value'),
      id: `${date}${month}${year}${reimbursementRef.id}`,
      key: momentNow.clone().startOf('day').valueOf(),
      currency: 'INR',
      reimbursementType: template,
      reimbursementName:
        locals.change.after.get('attachment.Claim Type.value') || '',
      details: {
        status,
        claimId: activityId,
        rate: null,
        checkInTimestamp: null,
        startLocation: null,
        endLocation: null,
        distanceTravelled: locals.addendumDocData.distanceTravelled,
        photoURL: locals.change.after.get('attachment.Photo URL.value') || '',
      },
    },
    { merge: true },
  );

  return batch.commit();
};

const reimburseDailyAllowance = async locals => {
  if (!locals.addendumDocData) {
    return;
  }

  if (
    locals.addendumDocData.isSupportRequest ||
    locals.addendumDocData.isAdminRequest
  ) {
    return;
  }

  const reimbursementType = 'daily allowance';
  const {
    officeId,
    office,
    timezone,
    creator: { phoneNumber },
  } = locals.change.after.data();
  const { timestamp } = locals.addendumDocData;
  const momentNow = momentTz(timestamp).tz(timezone);
  const action = locals.addendumDocData.action;
  const scheduledOnly = action === httpsActions.checkIn;
  const existingDailyAllowances = new Set();
  const batch = db.batch();
  const date = momentNow.date();
  const month = momentNow.month();
  const year = momentNow.year();
  let uid = locals.addendumDocData.uid;

  if (!uid) {
    uid = (await getAuth(phoneNumber)).uid;
  }

  const allowancesToday = await rootCollections.offices
    .doc(officeId)
    .collection(subcollectionNames.REIMBURSEMENTS)
    .where('date', '==', date)
    .where('month', '==', month)
    .where('year', '==', year)
    .where('reimbursementType', '==', reimbursementType)
    .where('phoneNumber', '==', phoneNumber)
    .get();

  /**
   * one type of reimbursement in a single day will
   * be given to a user only once.
   */
  allowancesToday.forEach(doc =>
    existingDailyAllowances.add(doc.get('reimbursementName')),
  );

  let dailyAllowanceBaseQuery = rootCollections.activities
    .where('template', '==', reimbursementType)
    .where('status', '==', 'CONFIRMED')
    .where('officeId', '==', officeId);

  if (scheduledOnly) {
    dailyAllowanceBaseQuery = dailyAllowanceBaseQuery.where(
      'attachment.Scheduled Only.value',
      '==',
      true,
    );
  }

  // action
  const dailyAllowanceActivities = await dailyAllowanceBaseQuery.get();

  dailyAllowanceActivities.forEach(daActivity => {
    const { value: attachmentName } = daActivity.get('attachment.Name');
    const [startHours, startMinutes] = daActivity
      .get('attachment.Start Time.value')
      .split(':');
    const [endHours, endMinutes] = daActivity
      .get('attachment.End Time.value')
      .split(':');

    if (
      startHours === '' ||
      startMinutes === '' ||
      endHours === '' ||
      endMinutes === ''
    ) {
      return;
    }

    const momentStart = momentNow
      .clone()
      .hours(startHours)
      .minutes(startMinutes);
    const momentEnd = momentNow.clone().hours(endHours).minutes(endMinutes);
    const includeStartAndEndTime = '[]';
    const isInRange = momentNow.isBetween(
      momentStart,
      momentEnd,
      'minutes',
      includeStartAndEndTime,
    );

    /** Is not in the time range */
    if (!isInRange) {
      return;
    }

    if (existingDailyAllowances.has(attachmentName)) {
      return;
    }

    const update = Object.assign(
      {},
      getRoleReportData(locals.roleObject, phoneNumber),
      {
        uid,
        date,
        month,
        year,
        officeId,
        phoneNumber,
        reimbursementType,
        currency: 'INR',
        timestamp: Date.now(),
        office: locals.change.after.get('office'),
        checkInTimestamp: locals.change.after.get('timestamp'),
        reimbursementName: daActivity.get('attachment.Name.value'),
        amount: daActivity.get('attachment.Amount.value'),
        relevantActivityId: locals.change.after.id,
        dailyAllowanceActivityId: daActivity.id,
        currentGeopoint: locals.addendumDocData.location,
        previousGeopoint: (() => {
          if (
            locals.previousAddendumDoc &&
            locals.previousAddendumDoc.get('location')
          ) {
            return locals.previousAddendumDoc.get('location');
          }

          return null;
        })(),
        currentIdentifier: (() => {
          if (locals.addendumDocData.venueQuery) {
            return locals.addendumDocData.venueQuery.location;
          }

          return locals.addendumDocData.identifier;
        })(),
        previousIdentifier: (() => {
          if (
            locals.previousAddendumDoc &&
            locals.previousAddendumDoc.get('venueQuery.location')
          ) {
            return locals.previousAddendumDoc.get('venueQuery.location');
          }

          if (
            locals.previousAddendumDoc &&
            locals.previousAddendumDoc.get('identifier')
          ) {
            return locals.previousAddendumDoc.get('identifier');
          }

          return null;
        })(),
      },
    );

    const ref = rootCollections.offices
      .doc(officeId)
      .collection(subcollectionNames.REIMBURSEMENTS)
      .doc();

    batch.set(ref, update, { merge: true });

    batch.set(
      rootCollections.updates
        .doc(uid)
        .collection(subcollectionNames.ADDENDUM)
        .doc(),
      Object.assign(
        {},
        {
          date,
          month,
          year,
          office,
          officeId,
          currency: 'INR',
          timestamp: Date.now(),
          amount: daActivity.get('attachment.Amount.value'),
          _type: addendumTypes.REIMBURSEMENT,
          key: momentNow.clone().startOf('day').valueOf(),
          id: `${date}${month}${year}${ref.id}`,
          reimbursementType: daActivity.get('template'),
          reimbursementName: daActivity.get('attachment.Name.value'),
          details: {
            rate: null,
            checkInTimestamp: momentNow.tz(timezone).valueOf(), // unix
            startLocation: null, // start
            endLocation: null, // end
            distanceTravelled: null,
            photoURL: null,
            status: null,
            claimId: null,
          },
        },
      ),
      { merge: true },
    );
  });

  return batch.commit();
};

const getStartPointObject = async ({
  startPointLatitude,
  startPointLongitude,
  baseLocation,
  officeId,
}) => {
  const identifier = 'Start Point';
  // start point present => use that
  if (
    typeof startPointLatitude === 'number' &&
    typeof startPointLongitude === 'number'
  ) {
    return {
      identifier,
      geopoint: {
        latitude: startPointLatitude,
        longitude: startPointLongitude,
      },
    };
  }

  if (!baseLocation) {
    return null;
  }

  // else check for base location
  // if base location is set
  // use that as start point
  const {
    docs: [baseLocationDoc],
  } = await rootCollections.activities
    .where('attachment.Name.value', '==', baseLocation)
    .where('officeId', '==', officeId)
    .where('status', '==', 'CONFIRMED')
    .limit(1)
    .get();

  if (!baseLocationDoc) {
    return null;
  }

  const [{ geopoint }] = baseLocationDoc.get('venue');

  return {
    identifier,
    geopoint: {
      latitude: geopoint._latitude || geopoint.latitude,
      longitude: geopoint._longitude || geopoint.longitude,
    },
  };
};

const getLatLngObject = ({ latLngObject }) => ({
  latitude: latLngObject.latitude || latLngObject._latitude,
  longitude: latLngObject.longitude || latLngObject._longitude,
});

const reimburseKmAllowance = async locals => {
  // if action is create, checkin - then look
  // for scheduled only false in
  // employee object and make km allowance if available
  // basis same logic of previous checkin is same
  // date then km allowance
  // between the two else km allowance from
  // startpoint/base location to both

  // if action is checkin - then look for scheduled
  // only true in employee and make km allowance
  // if available basis same logic of previous
  // action == checkin, same date then km
  // allowance between the two else km allowance
  // from start point/base location from both

  if (!locals.addendumDocData) {
    return;
  }

  // Support/admin requests are not eligible for reimbursements
  if (
    locals.addendumDocData.isSupportRequest ||
    locals.addendumDocData.isAdminRequest
  ) {
    return;
  }

  const { template, office, officeId, timezone } = locals.change.after.data();
  const { timestamp } = locals.addendumDocData;
  let uid = locals.addendumDocData.uid;

  // Doing this because action=check-in can show up with
  // non-check-in template activities too.
  const phoneNumber = (() => {
    if (template === 'check-in') {
      return locals.change.after.get('creator.phoneNumber');
    }

    return locals.addendumDocData.user;
  })();

  if (!uid) {
    uid = (await getAuth(phoneNumber)).uid;
  }

  const roleDoc = locals.roleObject;

  // Not an employee, km allowance is skipped
  if (!roleDoc) {
    return;
  }

  /**
   * Role might not have KM Rate since that field is currently in 'employee' activities
   */
  const { value: kmRate } = roleDoc.attachment['KM Rate'] || {};
  const { value: startPointLatitude } =
    roleDoc.attachment['Start Point Latitude'] || {};
  const { value: startPointLongitude } =
    roleDoc.attachment['Start Point Longitude'] || {};
  const { value: scheduledOnly } = roleDoc.attachment['Scheduled Only'] || {};

  // Scheduled Only means action === check-in. Exit otherwise
  if (scheduledOnly && locals.addendumDocData.action !== httpsActions.checkIn) {
    return;
  }

  if (!kmRate) {
    return;
  }

  const roleData = {
    phoneNumber,
    employeeName: roleDoc.attachment.Name.value,
    employeeCode: roleDoc.attachment['Employee Code'].value,
    baseLocation: roleDoc.attachment['Base Location'].value,
    region: roleDoc.attachment.Region.value,
    department: roleDoc.attachment.Department.value,
    minimumDailyActivityCount:
      roleDoc.attachment['Minimum Daily Activity Count'].value,
    minimumWorkingHours: roleDoc.attachment['Minimum Working Hours'].value,
  };

  const batch = db.batch();
  const reimbursementType = 'km allowance';
  const momentNow = momentTz(timestamp).tz(timezone);
  const date = momentNow.date();
  const month = momentNow.month();
  const year = momentNow.year();
  const commonReimObject = Object.assign(
    {},
    {
      date,
      month,
      year,
      office,
      uid,
      officeId,
      phoneNumber,
      reimbursementType,
      currency: 'INR',
      timestamp: Date.now(),
      relevantActivityId: locals.change.after.id,
      employeeName: roleDoc.attachment.Name.value,
    },
  );

  const [
    previousKmReimbursementQuery,
    previousReimbursementUpdateQuery,
  ] = await Promise.all([
    rootCollections.offices
      .doc(officeId)
      .collection(subcollectionNames.REIMBURSEMENTS)
      .where('intermediate', '==', true)
      .where('date', '==', date)
      .where('month', '==', month)
      .where('year', '==', year)
      .where('reimbursementType', '==', reimbursementType)
      .where('phoneNumber', '==', phoneNumber)
      .limit(1)
      .get(),
    rootCollections.updates
      .doc(uid)
      .collection(subcollectionNames.ADDENDUM)
      .where('_type', '==', addendumTypes.REIMBURSEMENT)
      .where('intermediate', '==', true)
      .where('date', '==', date)
      .where('month', '==', month)
      .where('year', '==', year)
      .where('reimbursementType', '==', reimbursementType)
      .limit(1)
      .get(),
  ]);

  const startPointDetails = await getStartPointObject({
    officeId,
    startPointLatitude,
    startPointLongitude,
    baseLocation: roleData.baseLocation,
  });

  if (!startPointDetails) {
    return;
  }

  const distanceBetweenCurrentAndStartPoint = await getDistanceFromDistanceMatrix(
    startPointDetails.geopoint,
    locals.addendumDocData.location,
  );

  if (distanceBetweenCurrentAndStartPoint < 1) {
    return;
  }

  /**
   * This function should be refactored. But I'm short on time.
   * Also don't want to touch it since it might break stuff.
   */
  if (previousKmReimbursementQuery.empty) {
    // create km allowance for start point to current location
    // create km allowance for current location to start point.
    const firstReimbursementRef = rootCollections.offices
      .doc(officeId)
      .collection(subcollectionNames.REIMBURSEMENTS)
      .doc();
    const secondReimbursementRef = rootCollections.offices
      .doc(officeId)
      .collection(subcollectionNames.REIMBURSEMENTS)
      .doc();
    const firstReimbursementUpdateRef = rootCollections.updates
      .doc(uid)
      .collection(subcollectionNames.ADDENDUM)
      .doc();
    const secondReimbursementUpdateRef = rootCollections.updates
      .doc(uid)
      .collection(subcollectionNames.ADDENDUM)
      .doc();

    /**
     * Amount earned for travelling between start point
     * to the current location.
     */
    const amountEarned = currencyJs(kmRate).multiply(
      distanceBetweenCurrentAndStartPoint,
    );

    const amountEarnedInString = amountEarned.toString();

    // startPoint (previous) to current location(current)
    batch.set(
      firstReimbursementRef,
      Object.assign({}, roleData, commonReimObject, {
        amount: amountEarnedInString,
        distance: distanceBetweenCurrentAndStartPoint,
        previousIdentifier: startPointDetails.identifier,
        previousGeopoint: startPointDetails.geopoint,
        currentIdentifier: (() => {
          if (locals.addendumDocData.venueQuery) {
            return locals.addendumDocData.venueQuery.location;
          }

          return locals.addendumDocData.identifier;
        })(),
        currentGeopoint: getLatLngObject({
          latLngObject: locals.addendumDocData.location,
        }),
        intermediate: false,
      }),
      { merge: true },
    );

    // current location to start point
    batch.set(
      secondReimbursementRef,
      Object.assign({}, commonReimObject, {
        rate: kmRate,
        previousIdentifier: (() => {
          if (locals.addendumDocData.venueQuery) {
            return locals.addendumDocData.venueQuery.location;
          }

          return locals.addendumDocData.identifier;
        })(),
        previousGeopoint: getLatLngObject({
          latLngObject: locals.addendumDocData.location,
        }),
        currentIdentifier: startPointDetails.identifier,
        currentGeopoint: startPointDetails.geopoint,
        intermediate: true,
        amount: amountEarnedInString,
        distance: distanceBetweenCurrentAndStartPoint,
      }),
    );

    // start point to current location
    batch.set(
      firstReimbursementUpdateRef,
      Object.assign({}, commonReimObject, {
        amount: amountEarnedInString,
        _type: addendumTypes.REIMBURSEMENT,
        id: `${date}${month}${year}${firstReimbursementRef.id}`,
        key: momentNow.clone().startOf('day').valueOf(),
        reimbursementName: null,
        details: {
          rate: kmRate,
          startLocation: startPointDetails.geopoint,
          checkInTimestamp: locals.change.after.get('timestamp'),
          endLocation: getLatLngObject({
            latLngObject: locals.addendumDocData.location,
          }),
          distanceTravelled: distanceBetweenCurrentAndStartPoint,
          photoURL: null,
          status: null,
          claimId: null,
        },
      }),
    );

    // curr to start point
    batch.set(
      secondReimbursementUpdateRef,
      Object.assign({}, commonReimObject, {
        _type: addendumTypes.REIMBURSEMENT,
        amount: amountEarnedInString,
        id: `${date}${month}${year}${secondReimbursementRef.id}`,
        key: momentNow.clone().startOf('day').valueOf(),
        reimbursementName: null,
        intermediate: true,
        details: {
          rate: kmRate,
          startLocation: getLatLngObject({
            latLngObject: locals.addendumDocData.location,
          }),
          checkInTimestamp: locals.change.after.get('timestamp'),
          endLocation: startPointDetails.geopoint,
          distanceTravelled: distanceBetweenCurrentAndStartPoint,
          photoURL: null,
          status: null,
          claimId: null,
        },
      }),
    );
  } else {
    if (locals.addendumDocData.distanceTravelled < 1) {
      return;
    }

    const amountThisTime = currencyJs(kmRate)
      .multiply(locals.addendumDocData.distanceTravelled)
      .toString();
    const {
      docs: [oldReimbursementDoc],
    } = previousKmReimbursementQuery;
    const {
      docs: [oldUpdatesDoc],
    } = previousReimbursementUpdateQuery;
    const r1 = rootCollections.offices
      .doc(officeId)
      .collection(subcollectionNames.REIMBURSEMENTS)
      .doc();
    const u1 = rootCollections.updates
      .doc(uid)
      .collection(subcollectionNames.ADDENDUM)
      .doc();

    // r2
    batch.set(
      oldReimbursementDoc.ref,
      Object.assign({}, roleData, commonReimObject, {
        rate: kmRate,
        amount: amountThisTime,
        intermediate: false,
        currentIdentifier: (() => {
          if (locals.addendumDocData.venueQuery) {
            return locals.addendumDocData.venueQuery.location;
          }

          return locals.addendumDocData.identifier;
        })(),
        currentGeopoint: getLatLngObject({
          latLngObject: locals.addendumDocData.location,
        }),
      }),
      { merge: true },
    );

    const oldUpdatesRef = (() => {
      if (oldUpdatesDoc) {
        return oldUpdatesDoc.ref;
      }

      return rootCollections.updates
        .doc(uid)
        .collection(subcollectionNames.ADDENDUM)
        .doc();
    })();

    batch.set(
      oldUpdatesRef,
      Object.assign({}, commonReimObject, {
        date,
        month,
        year,
        id: `${date}${month}${year}${oldReimbursementDoc.id}`,
        key: momentTz()
          .date(date)
          .month(month)
          .year(year)
          .startOf('date')
          .valueOf(),
        amount: amountThisTime,
        _type: addendumTypes.REIMBURSEMENT,
        reimbursementName: null,
        intermediate: false,
        details: {
          rate: kmRate,
          startLocation: oldReimbursementDoc.get('previousGeopoint'),
          distanceTravelled: locals.addendumDocData.distanceTravelled,
          photoURL: null,
          status: null,
          claimId: null,
          checkInTimestamp: timestamp,
          endLocation: getLatLngObject({
            latLngObject: locals.addendumDocData.location,
          }),
        },
      }),
      { merge: true },
    );

    // currentLocation (start) to startPoint (end)
    batch.set(
      r1,
      Object.assign({}, roleData, commonReimObject, {
        // cumulativeAmount,
        rate: kmRate,
        amount: amountThisTime,
        currentIdentifier: startPointDetails.identifier,
        currentGeopoint: startPointDetails.geopoint,
        previousIdentifier: (() => {
          if (locals.addendumDocData.venueQuery) {
            return locals.addendumDocData.venueQuery.location;
          }

          return locals.addendumDocData.identifier;
        })(),
        previousGeopoint: getLatLngObject({
          latLngObject: locals.addendumDocData.location,
        }),
        intermediate: true,
      }),
      { merge: true },
    );

    // currentLocation (start) to startPoint (end)
    batch.set(
      u1,
      Object.assign({}, commonReimObject, {
        _type: addendumTypes.REIMBURSEMENT,
        amount: amountThisTime,
        id: `${date}${month}${year}${r1.id}`,
        key: momentTz()
          .date(date)
          .month(month)
          .year(year)
          .startOf('date')
          .valueOf(),
        reimbursementName: null,
        intermediate: true,
        details: {
          rate: kmRate,
          startLocation: getLatLngObject({
            latLngObject: locals.addendumDocData.location,
          }),
          checkInTimestamp: locals.change.after.get('timestamp'),
          endLocation: startPointDetails.geopoint,
          distanceTravelled: distanceBetweenCurrentAndStartPoint,
          photoURL: null,
          status: null,
          claimId: null,
        },
      }),
    );
  }

  return batch.commit();
};

const handleReimbursement = async locals => {
  if (!locals.addendumDocData) {
    return;
  }

  /** Support creates/updates stuff */
  if (
    locals.addendumDocData.isSupportRequest ||
    locals.addendumDocData.isAdminRequest
  ) {
    return;
  }

  const { template } = locals.change.after.data();

  if (template === 'claim') {
    return reimburseClaim(locals);
  }

  return;
};

const getLateStatus = ({ firstCheckInTimestamp, dailyStartTime, timezone }) => {
  if (Number.isInteger(firstCheckInTimestamp)) {
    return false;
  }

  if (!isNonEmptyString(dailyStartTime)) {
    return false;
  }

  const [startHours, startMinutes] = dailyStartTime.split(':');
  const momentNow = momentTz(firstCheckInTimestamp).tz(timezone);
  const momentStartTime = momentTz().hour(startHours).minutes(startMinutes);

  return momentNow.diff(momentStartTime, 'minutes', true) > 15;
};

/**
 *
 * @param {Set} leaveSnaps Returns a set of leave dates using
 * the snapshot in the format `2nd Jan 2020`.
 */
const getAllScheduleDatesFromActivitySnaps = leaveSnaps => {
  const leaveDates = new Set();
  leaveSnaps.forEach(leaves => {
    /**
     * @param {{ data: () => { scheduleDates: any; status: any; }; }} leave
     */
    leaves.forEach(leave => {
      const { scheduleDates, status } = leave.data();

      if (status === 'CANCELLED') {
        return;
      }

      /**
       * Some old leave activities might not have the scheduleDates array
       * But still, this `OR` clause is redundant in 99.9% of ActivityOnWrite instances.
       */
      (scheduleDates || []).forEach(dateItem => leaveDates.add(dateItem));
    });
  });

  return leaveDates;
};

const newBackfill = async locals => {
  // if prevCheckIn is in current month
  // return

  // fetch base location of this user
  // fetch leaves of this user with scheduleDates array_contains
  // of previous month dates

  // fetch prev month attendance doc
  // if doc exists
  // range => (prev checkIn date to today)

  // if doc doesn't exist
  // rangeStart = maxUnix(1st of prevMonth, empCreationUnix, lastCheckInTs)
  // rangeEnd = end of last month
  const {
    /**
     * This field could be undefined, which will make moment return current
     * timestamp with current month
     */
    lastTimestamp,
    activityData,
    uid,
    user: phoneNumber,
  } = locals.addendumDocData;
  const { timezone, office, officeId } = activityData;
  const prevCheckInMoment = momentTz(lastTimestamp).tz(timezone);
  const momentNow = momentTz().tz(timezone);
  const momentPrevMonth = momentNow.clone().subtract(1, 'month');

  // if lastTimestamp is a unix timestamp of current month
  // momentTz will fallback to the current timestamp.
  // This clause will run, so, our logic is fine.
  if (prevCheckInMoment.month() === momentNow.month()) {
    return;
  }

  if (!locals.roleObject) {
    return;
  }

  console.log('backfill called');

  // fetch base location
  // fetch leaves with scheduleDates in previousMonth
  const {
    docs: [attendanceDocPrevMonth],
  } = await rootCollections.offices
    .doc(officeId)
    .collection(subcollectionNames.ATTENDANCES)
    .where('month', '==', momentPrevMonth.month())
    .where('year', '==', momentPrevMonth.year())
    .where('phoneNumber', '==', phoneNumber)
    .limit(1)
    .get();

  const { value: baseLocation } =
    locals.roleObject.attachment['Base Location'] || {};

  const {
    docs: [branchDoc],
  } = await rootCollections.activities
    .where('template', '==', 'branch')
    .where('status', '==', 'CONFIRMED')
    .where('officeId', '==', officeId)
    .where('attachment.Name.value', '==', baseLocation)
    .limit(1)
    .get();

  const getRangeStart = () => {
    if (attendanceDocPrevMonth && lastTimestamp) {
      return momentTz(lastTimestamp).tz(timezone);
    }

    const max = Math.max(
      momentPrevMonth.clone().startOf('month').valueOf(),
      locals.roleObject.createTime,
      momentTz(lastTimestamp).tz(timezone).valueOf(),
    );

    return momentTz(max).tz(timezone).startOf('date');
  };

  const leavePromises = [];
  const iterationRangeEnd = momentPrevMonth.clone().endOf('month');
  const iterationRangeStart = getRangeStart();
  const momentInterator = iterationRangeStart.clone();
  const attendanceData = attendanceDocPrevMonth
    ? attendanceDocPrevMonth.data()
    : {};

  attendanceData.attendance = attendanceData.attendance || {};

  if (branchDoc) {
    const { value: weeklyOffFromBranch } = branchDoc.get(
      'attachment.Weekly Off',
    );

    while (momentInterator.isSameOrBefore(iterationRangeEnd)) {
      const { date } = momentInterator.toObject();
      const weekdayName = momentInterator
        .format(dateFormats.WEEKDAY)
        .toLowerCase();
      const weeklyOff =
        weeklyOffFromBranch.toLowerCase() === weekdayName.toLowerCase();
      attendanceData.attendance[date] =
        attendanceData.attendance[date] || getDefaultAttendanceObject();
      attendanceData.attendance[date].weeklyOff = weeklyOff;

      if (weeklyOff) {
        attendanceData.attendance[date].attendance = 1;
      }

      leavePromises.push(
        rootCollections.activities
          .where('creator.phoneNumber', '==', phoneNumber)
          .where('template', '==', 'leave')
          .where('officeId', '==', officeId)
          .where('status', 'in', ['CONFIRMED', 'PENDING'])
          .where(
            'scheduleDates',
            'array-contains',
            momentInterator.format(dateFormats.DATE),
          )
          .get(),
      );

      momentInterator.add(1, 'day');
    }
  }

  const secondIterator = iterationRangeStart.clone();
  const leaveSnaps = await Promise.all(leavePromises);
  const leaveDates = getAllScheduleDatesFromActivitySnaps(leaveSnaps);

  while (secondIterator.isSameOrBefore(iterationRangeEnd)) {
    const { date } = secondIterator.toObject();
    const onLeave = leaveDates.has(secondIterator.format(dateFormats.DATE));

    attendanceData.attendance[date] =
      attendanceData.attendance[date] || getDefaultAttendanceObject();
    attendanceData.attendance[date].onLeave = onLeave;

    if (onLeave) {
      attendanceData.attendance[date].attendance = 1;
    }

    secondIterator.add(1, 'day');
  }

  const ref = attendanceDocPrevMonth
    ? attendanceDocPrevMonth.ref
    : rootCollections.offices
        .doc(officeId)
        .collection(subcollectionNames.ATTENDANCES)
        .doc();

  return ref.set(
    Object.assign(
      {},
      getRoleReportData(locals.roleObject, phoneNumber),
      attendanceData,
      {
        uid,
        officeId,
        phoneNumber,
        office,
        month: momentPrevMonth.month(),
        year: momentPrevMonth.year(),
        roleDoc: locals.roleObject || null,
      },
    ),
    { merge: true },
  );
};

const getWeeklyOffDatesInMonth = ({
  weeklyOff,
  attendanceMonth,
  attendanceYear,
}) => {
  const result = [];
  const momentInstance = momentTz().month(attendanceMonth).year(attendanceYear);
  const monthStart = momentInstance.clone().startOf('month');
  const monthEnd = monthStart.clone().endOf('month');
  const iterator = monthStart.clone();

  while (iterator.isSameOrBefore(monthEnd)) {
    const { date } = iterator.toObject();

    if (
      weeklyOff.toLowerCase() ===
      iterator.format(dateFormats.WEEKDAY).toLowerCase()
    ) {
      result.push(date);
    }

    iterator.add(1, 'day');
  }

  return result;
};

const populateMissingAttendancesInCurrentMonth = async (
  locals,
  attendanceUpdate,
  attendanceDocRef,
  attendanceDocExistsAlready,
) => {
  // fetch this user's leaves, branch
  // populate weekly off, holidays and weeklyOffs
  // Populate in this month's attendance doc

  // Attendance doc was not created in this instance, so there is no
  // point of putting branch/leave etc every time.
  if (attendanceDocExistsAlready) {
    return;
  }

  const batch = db.batch();
  const { month: attendanceMonth, year: attendanceYear } = attendanceUpdate;

  const {
    officeId,
    timezone = 'Asia/Kolkata',
    creator: { phoneNumber },
  } = locals.change.after.data();
  const momentInstance = momentTz()
    .month(attendanceMonth)
    .year(attendanceYear)
    .tz(timezone);
  const monthStart = momentInstance.clone().startOf('month');
  const monthEnd = monthStart.clone().endOf('month');
  const iterator = monthStart.clone();
  const leavePromises = [];

  while (iterator.isSameOrBefore(monthEnd)) {
    leavePromises.push(
      rootCollections.profiles
        .doc(phoneNumber)
        .collection(subcollectionNames.ACTIVITIES)
        .where('creator.phoneNumber', '==', phoneNumber)
        .where('officeId', '==', officeId)
        .where('template', '==', 'leave')
        .where('status', 'in', ['CONFIRMED', 'PENDING'])
        .where(
          'scheduleDates',
          'array-contains',
          iterator.format(dateFormats.DATE),
        )
        .get(),
    );

    iterator.add(1, 'day');
  }

  const leaveIdUniques = new Set();
  const leavesThisMonth = [];

  (await Promise.all(leavePromises)).forEach(leaves => {
    leaves.forEach(leave => {
      const { id: leaveId } = leave;
      const [{ startTime, endTime }] = leave.get('schedule');
      const isRepeat = leaveIdUniques.has(leaveId);
      const {
        attachment: { 'Leave Type': { value } = {} } = {},
      } = leave.data();

      // we are using array-contains which can return the same leave activity
      // if the activity has more than 1 schedule date.
      if (isRepeat) {
        return;
      }

      leaveIdUniques.add(leaveId);

      if (
        momentTz(startTime).isBetween(monthStart, monthEnd) ||
        momentTz(endTime).isBetween(monthStart, monthEnd)
      ) {
        // these activities have schedule in the current month
        leavesThisMonth.push({ leaveType: value, startTime, endTime });
      }
    });
  });

  leavesThisMonth.forEach(({ leaveType = '', startTime, endTime }) => {
    const leaveStartMoment = momentTz(startTime);
    const leaveEndMoment = momentTz(endTime);
    const iteratorMoment = leaveStartMoment.clone();

    while (iteratorMoment.isSameOrBefore(leaveEndMoment)) {
      const { date, months: month, years: year } = iteratorMoment.toObject();

      // If we don't check this, then the loop will overwrite
      // this month's dates without consideration for the
      // actual month/year
      if (attendanceMonth !== month || attendanceYear !== year) {
        continue;
      }

      attendanceUpdate.attendance = attendanceUpdate.attendance || {};
      attendanceUpdate.attendance[date] =
        attendanceUpdate.attendance[date] || getDefaultAttendanceObject();
      attendanceUpdate.attendance[date].attendance = 1;

      attendanceUpdate.attendance[date].onLeave = true;
      attendanceUpdate.attendance[date].leaveType = leaveType;

      iteratorMoment.add(1, 'day');
    }
  });

  const baseLocation =
    locals.roleObject &&
    locals.roleObject.attachment['Base Location'] &&
    locals.roleObject.attachment['Base Location'].value;

  if (baseLocation) {
    const {
      docs: [branchDoc],
    } = await rootCollections.activities
      .where('officeId', '==', officeId)
      .where(
        'template',
        '==',
        locals.roleObject.attachment['Base Location'].type,
      )
      .where('status', '==', 'CONFIRMED')
      .where('attachment.Name.value', '==', baseLocation)
      .limit(1)
      .get();

    if (branchDoc) {
      const weeklyOff = branchDoc.get('attachment.Weekly Off.value');
      const weeklyOffThisMonth = getWeeklyOffDatesInMonth({
        weeklyOff,
        attendanceMonth,
        attendanceYear,
      });

      weeklyOffThisMonth.forEach(date => {
        attendanceUpdate.attendance = attendanceUpdate.attendance || {};
        attendanceUpdate.attendance[date] =
          attendanceUpdate.attendance[date] || getDefaultAttendanceObject();
        attendanceUpdate.attendance[date].attendance = 1;
        attendanceUpdate.attendance[date].weeklyOff = true;
      });

      branchDoc.get('schedule').forEach(({ startTime }) => {
        if (typeof startTime === 'number') {
          // stuff
          const { date, months: month, years: year } = momentTz(startTime)
            .tz(timezone)
            .toObject();

          if (attendanceMonth === month && attendanceYear === year) {
            attendanceUpdate.attendance = attendanceUpdate.attendance || {};
            attendanceUpdate.attendance[date] =
              attendanceUpdate.attendance[date] || getDefaultAttendanceObject();
            attendanceUpdate.attendance[date].attendance = 1;
            attendanceUpdate.attendance[date].holiday = true;
          }
        }
      });
    }
  }

  batch.set(attendanceDocRef, attendanceUpdate, { merge: true });

  return batch.commit();
};

const handleWorkday = async locals => {
  /**
   * Template === check-in and action === 'create'
   */
  if (!locals.addendumDocData) {
    return;
  }

  if (locals.addendumDocData.action !== httpsActions.create) {
    return;
  }

  const {
    officeId,
    timezone,
    office,
    creator: { phoneNumber },
  } = locals.change.after.data();
  const momentNow = momentTz(locals.addendumDocData.timestamp).tz(timezone);
  const todaysDate = momentNow.date();
  const month = momentNow.month();
  const year = momentNow.year();
  const roleData = getRoleReportData(locals.roleObject, phoneNumber);

  // If employee Location Validation Check => true
  // AND distanceAccurate => false
  // skip
  // Using explicit check for this case because
  // values can be empty strings.
  if (
    locals.roleObject &&
    locals.roleObject.attachment['Location Validation Check'].value === true &&
    locals.addendumDocData.distanceAccurate === false
  ) {
    return;
  }

  const { location } = locals.addendumDocData;
  const batch = db.batch();
  let uid = locals.addendumDocData.uid;

  if (!uid) {
    uid = (await getAuth(phoneNumber)).uid;
  }

  /**
   * This query might return 0 docs if the date = 1 in the month
   * or the user hasn't done anything since the start of the month
   */
  const {
    docs: [attendanceDoc],
  } = await rootCollections.offices
    .doc(officeId)
    .collection(subcollectionNames.ATTENDANCES)
    .where('phoneNumber', '==', phoneNumber)
    .where('month', '==', month)
    .where('year', '==', year)
    .limit(1)
    .get();

  const attendanceObject = attendanceDoc ? attendanceDoc.data() : {};

  attendanceObject.attendance = attendanceObject.attendance || {};
  attendanceObject.attendance[todaysDate] =
    attendanceObject.attendance[todaysDate] || getDefaultAttendanceObject();
  attendanceObject.attendance[todaysDate].working =
    attendanceObject.attendance[todaysDate].working || {};

  /**
   * If the first check-in has already been set for this user
   * we don't need to update it again for the day
   */
  attendanceObject.attendance[todaysDate].working.firstCheckInTimestamp =
    attendanceObject.attendance[todaysDate].working.firstCheckInTimestamp ||
    locals.addendumDocData.timestamp;

  attendanceObject.attendance[todaysDate].working.lastCheckInTimestamp =
    locals.addendumDocData.timestamp;
  attendanceObject.attendance[todaysDate].isLate = getLateStatus({
    timezone,
    firstCheckInTimestamp:
      attendanceObject.attendance[todaysDate].working.firstCheckInTimestamp,
    dailyStartTime: roleData.dailyStartTime,
  });

  attendanceObject.attendance[todaysDate].addendum =
    attendanceObject.attendance[todaysDate].addendum || [];

  // If we trigger activityOnWrite for check-in, the same addendum object
  // should not be pushed to the array since that will result in duplication
  // and the count of addendum will also be off.
  // This line handles the idempotency.
  const idx = attendanceObject.attendance[todaysDate].addendum.findIndex(
    val => {
      return val.addendumId === locals.addendumDoc.id;
    },
  );

  if (idx === -1) {
    attendanceObject.attendance[todaysDate].addendum.push({
      timestamp: locals.addendumDocData.timestamp,
      latitude: location._latitude || location.latitude,
      longitude: location._longitude || location.longitude,
      addendumId: locals.addendumDoc.id,
    });
  }

  /**
   * Sometimes when the code crashes or when an event is missed
   * we trigger `activityOnWrite` by updating the timestamp.
   * In that case, the sorting of the timestamps in this array
   * might get messed up. Sorting regardless helps us migitate
   * this case.
   */
  attendanceObject.attendance[todaysDate].addendum.sort(
    (first, second) => first.timestamp - second.timestamp,
  );

  const { length: numberOfCheckIns } = attendanceObject.attendance[
    todaysDate
  ].addendum;
  const [firstAddendum] = attendanceObject.attendance[todaysDate].addendum;
  const lastAddendum =
    attendanceObject.attendance[todaysDate].addendum[numberOfCheckIns - 1];
  const hoursWorked = momentTz(lastAddendum.timestamp).diff(
    momentTz(firstAddendum.timestamp),
    'hours',
    true,
  );

  attendanceObject.attendance[todaysDate].attendance = getStatusForDay({
    // difference between first and last action in hours
    hoursWorked,
    // number of actions done in the day by the user
    numberOfCheckIns,
    minimumDailyActivityCount: roleData.minimumDailyActivityCount,
    minimumWorkingHours: roleData.minimumWorkingHours,
  });

  if (
    attendanceObject.attendance[todaysDate].onAr ||
    attendanceObject.attendance[todaysDate].onLeave ||
    attendanceObject.attendance[todaysDate].holiday ||
    attendanceObject.attendance[todaysDate].weeklyOff
  ) {
    attendanceObject.attendance[todaysDate].attendance = 1;
  }

  attendanceObject.attendance[todaysDate].working.numberOfCheckIns =
    attendanceObject.attendance[todaysDate].addendum.length;

  // Will iterate over 1 to (todays date - 1)
  // Eg. if today is 18 => range {1, 17}
  const tempDates = getNumbersbetween(1, todaysDate);
  tempDates.forEach(date => {
    // TODO: Dates in this loop should only be the ones where the user
    // was an employee
    attendanceObject.attendance[date] =
      attendanceObject.attendance[date] || getDefaultAttendanceObject();

    const checkInsArray = attendanceObject.attendance[date].addendum || [];
    const { length: numberOfCheckIns } = checkInsArray || [];

    const hoursWorked = (() => {
      if (checkInsArray.length === 0) {
        return 0;
      }

      const [firstCheckInAddendum] = checkInsArray;
      const lastCheckInAddendum = checkInsArray[checkInsArray.length - 1];
      const { timestamp: firstCheckInTimestamp } = firstCheckInAddendum;
      const { timestamp: lastCheckInTimestamp } = lastCheckInAddendum;
      const firstMoment = momentTz(firstCheckInTimestamp).tz(timezone);
      const lastMoment = momentTz(lastCheckInTimestamp).tz(timezone);

      return lastMoment.diff(firstMoment, 'hours', true);
    })();

    attendanceObject.attendance[date].attendance = (() => {
      const { onAr, onLeave, holiday, weeklyOff } = attendanceObject.attendance[
        date
      ];

      if (onAr || onLeave || holiday || weeklyOff) {
        return 1;
      }

      return getStatusForDay({
        hoursWorked,
        numberOfCheckIns,
        minimumDailyActivityCount: roleData.minimumDailyActivityCount,
        minimumWorkingHours: roleData.minimumWorkingHours,
      });
    })();

    batch.set(
      rootCollections.updates
        .doc(uid)
        .collection(subcollectionNames.ADDENDUM)
        .doc(),
      Object.assign({}, attendanceObject.attendance[date], {
        date,
        month,
        year,
        office,
        officeId,
        phoneNumber,
        timestamp: Date.now(),
        _type: addendumTypes.ATTENDANCE,
        id: `${date}${month}${year}${officeId}`,
        key: momentTz()
          .date(date)
          .month(month)
          .year(year)
          .startOf('date')
          .valueOf(),
      }),
    );
  });

  const attendanceDocRef = attendanceDoc
    ? attendanceDoc.ref
    : rootCollections.offices
        .doc(officeId)
        .collection(subcollectionNames.ATTENDANCES)
        .doc();

  const attendanceUpdate = Object.assign({}, roleData, attendanceObject, {
    uid,
    year,
    month,
    office,
    officeId,
    phoneNumber,
    roleDoc: locals.roleObject || null,
    timestamp: Date.now(),
  });

  batch.set(attendanceDocRef, attendanceUpdate, { merge: true });

  batch.set(
    rootCollections.updates
      .doc(uid)
      .collection(subcollectionNames.ADDENDUM)
      .doc(),
    Object.assign({}, attendanceObject.attendance[todaysDate], {
      month,
      year,
      office,
      officeId,
      phoneNumber,
      date: todaysDate,
      timestamp: Date.now(),
      _type: addendumTypes.ATTENDANCE,
      id: `${todaysDate}${month}${year}${officeId}`,
      key: momentNow.clone().startOf('date').valueOf(),
    }),
    { merge: true },
  );

  await batch.commit();

  await handleScheduledActivities(locals);

  /**
   * If the attendance doc is created, this function will fetch this user's
   * leaves, and holidays/weekly off (using branch).
   */
  await populateMissingAttendancesInCurrentMonth(
    locals,
    attendanceUpdate,
    attendanceDocRef,
    !!attendanceDoc,
  );

  /** Only populate the missing attendances when the attendance doc was created */
  if (attendanceDoc) {
    return;
  }

  if (!locals.roleObject) {
    return;
  }

  return newBackfill(locals);
};

const getPermutations = officeName => {
  const nameCombinations = new Set();
  const lowerCaseName = officeName.toLowerCase();

  [' ', '.', ',', '-', '&', '(', ')'].forEach(character => {
    lowerCaseName.split(character).forEach(part => nameCombinations.add(part));
  });

  return [...nameCombinations].filter(Boolean);
};

const getSubcollectionActivityObject = ({ activityObject, customerObject }) => {
  const { status, template, timezone = 'Asia/Kolkata' } = activityObject.data();
  const creationTimestamp = activityObject.createTime.toDate().getTime();
  const momentOfCreation = momentTz(creationTimestamp).tz(timezone);
  const intermediate = Object.assign({}, activityObject.data(), {
    creationTimestamp,
    timestamp: Date.now(),
    addendumDocRef: null,
    isCancelled: status === 'CANCELLED',
    creationDate: momentOfCreation.date(),
    creationMonth: momentOfCreation.month(),
    creationYear: momentOfCreation.year(),
  });

  if (template === 'office') {
    intermediate.slug = slugify(activityObject.get('attachment.Name.value'));
    intermediate.searchables = getPermutations(
      activityObject.get('attachment.Name.value'),
    );
  }

  if (customerObject) {
    intermediate.customerObject = customerObject;
  }

  return intermediate;
};

const attendanceHandler = async locals => {
  const {
    template,
    status,
    officeId,
    office,
    schedule,
    attachment,
    creator: { phoneNumber },
  } = locals.change.after.data();
  const {
    Reason: { value: reason },
  } = attachment;
  const isCancelled = status === 'CANCELLED';
  const [{ startTime, endTime }] = schedule;

  const roleData = getRoleReportData(locals.roleObject, phoneNumber);
  // uid should be of the creator in case of attendance
  // do not use uid from addendum.
  // that will introduce a bug where the Updates/{uid}/Addendum/{autoId}
  // doc will be created for the person changing the status
  const { uid } = await getAuth(phoneNumber);
  const batch = db.batch();
  const datesToSet = new Set();
  const momentStartTime = momentTz(startTime).startOf('date');
  const copyOfMomentStartTime = momentStartTime.clone();
  const momentEndTime = momentTz(endTime).endOf('date');
  const startTimeMonth = momentStartTime.month();
  const { user } = locals.addendumDocData || {};

  /**
   * Activity was cancelled during this instance
   */
  const hasBeenCancelled =
    locals.change.before.data() &&
    locals.change.before.get('status') !== 'CANCELLED' &&
    locals.change.after.get('status') === 'CANCELLED';

  while (copyOfMomentStartTime.isSameOrBefore(momentEndTime)) {
    if (startTimeMonth === copyOfMomentStartTime.month()) {
      datesToSet.add(copyOfMomentStartTime.date());
    }

    copyOfMomentStartTime.add(1, 'days');
  }

  const {
    docs: [attendanceDoc],
  } = await rootCollections.offices
    .doc(officeId)
    .collection(subcollectionNames.ATTENDANCES)
    .where('phoneNumber', '==', phoneNumber)
    .where('month', '==', momentStartTime.month())
    .where('year', '==', momentStartTime.year())
    .limit(1)
    .get();

  const attendanceData = attendanceDoc ? attendanceDoc.data() : {};

  attendanceData.attendance = attendanceData.attendance || {};
  datesToSet.forEach(date => {
    attendanceData.attendance[date] =
      attendanceData.attendance[date] || getDefaultAttendanceObject();

    if (!isCancelled) {
      attendanceData.attendance[date].attendance = 1;
    }

    if (template === 'leave') {
      attendanceData.attendance[date].leave.leaveType =
        attachment['Leave Type'].value;
      attendanceData.attendance[date].leave.reason = reason;
      attendanceData.attendance[date].onLeave = !isCancelled;

      if (user) {
        attendanceData.attendance[date].leave[status] = {
          phoneNumber: user,
          timestamp: Date.now(),
        };
      }
    }

    if (template === 'attendance regularization') {
      attendanceData.attendance[date].ar.reason = reason;
      attendanceData.attendance[date].onAr = !isCancelled;

      if (user) {
        attendanceData.attendance[date].ar[status] = {
          phoneNumber: user,
          timestamp: Date.now(),
        };
      }
    }

    if (hasBeenCancelled) {
      // recalculate attendance for these dates
      const checkInsArray = attendanceData.attendance[date].addendum || [];
      const { length: numberOfCheckIns } = checkInsArray;
      const hoursWorked = getAttendanceHoursWorked({ attendanceData, date });

      attendanceData.attendance[date].attendance = (() => {
        const { holiday, weeklyOff } = attendanceData.attendance[date];

        if (holiday || weeklyOff) {
          return 1;
        }

        return getStatusForDay({
          // number of actions done in the day by the user
          numberOfCheckIns,
          // difference between first and last action in hours,
          hoursWorked,
          minimumWorkingHours: roleData.minimumWorkingHours,
          minimumDailyActivityCount: roleData.minimumDailyActivityCount,
        });
      })();
    }

    // During phone number change, the uid of new number might not exist.
    if (uid) {
      batch.set(
        rootCollections.updates
          .doc(uid)
          .collection(subcollectionNames.ADDENDUM)
          .doc(),
        Object.assign({}, attendanceData.attendance[date], {
          date,
          office,
          officeId,
          timestamp: Date.now(),
          activityId: locals.change.after.id,
          _type: addendumTypes.ATTENDANCE,
          month: momentStartTime.month(),
          year: momentStartTime.year(),
          id: `${date}${momentStartTime.month()}${momentStartTime.year()}${officeId}`,
          key: momentStartTime.clone().date(date).startOf('day').valueOf(),
        }),
        { merge: true },
      );
    }
  });

  const attendanceRef = attendanceDoc
    ? attendanceDoc.ref
    : rootCollections.offices
        .doc(officeId)
        .collection(subcollectionNames.ATTENDANCES)
        .doc();

  batch.set(
    attendanceRef,
    Object.assign({}, attendanceData, roleData, {
      uid,
      office,
      officeId,
      month: momentStartTime.month(),
      year: momentStartTime.year(),
      roleDoc: locals.roleObject || null,
    }),
    { merge: true },
  );

  return batch.commit();
};

const handleRoleDocCancelled = async (
  activityNewData,
  batch,
  template,
  activityId,
) => {
  if (template === 'admin') {
    // if admin is cancelled, remove from custom claims
    const userRecord = await getAuth(
      activityNewData.attachment['Phone Number'].value,
    );
    const customClaims = Object.assign({}, userRecord.customClaims);
    const adminClaimsSet = new Set(customClaims.admin || []);
    adminClaimsSet.delete(activityNewData.office);
    customClaims.admin = Array.from(adminClaimsSet);
    return auth
      .setCustomUserClaims(userRecord.uid, customClaims)
      .catch(console.error);
  }
  if (template === 'employee') {
    // if employee is cancelled, remove his roledoc from roleReferences
    batch.set(
      rootCollections.profiles.doc(
        activityNewData.attachment['Phone Number'].value,
      ),
      {
        roleReferences: {
          [activityNewData.office]: admin.firestore.FieldValue.delete(),
        },
      },
      {
        merge: true,
      },
    );
  }
  if (template === 'subscription') {
    // if his subscription is cancelled, remove it from profiles as it creates burden on read api and is garbage
    batch.delete(
      rootCollections.profiles
        .doc(activityNewData.attachment['Phone Number'].value)
        .collection(subcollectionNames.SUBSCRIPTIONS)
        .doc(activityId),
    );
  }
};

const handleRoleDocConfirmed = async (
  activityNewData,
  batch,
  template,
  activityId,
) => {
  if (template === 'admin') {
    // if admin was confirmed or created , add to his custom claims
    const userRecord = await getAuth(
      activityNewData.attachment['Phone Number'].value,
    );
    const customClaims = Object.assign({}, userRecord.customClaims);
    const adminClaims = new Set(customClaims.admin).add(activityNewData.office);
    customClaims.admin = Array.from(adminClaims);
    return auth
      .setCustomUserClaims(userRecord.uid, customClaims)
      .catch(console.error);
  }
  if (template === 'employee') {
    // if employee is CONFIRMED, add him to roleReferences
    batch.set(
      rootCollections.profiles.doc(
        activityNewData.attachment['Phone Number'].value,
      ),
      {
        roleReferences: {
          [activityNewData.office]: activityNewData,
        },
      },
      {
        merge: true,
      },
    );
  }
  if (template === 'subscription') {
    // create his subscription under /profiles/{}/sub-collection_subscription/
    batch.set(
      rootCollections.profiles
        .doc(activityNewData.attachment['Phone Number'].value)
        .collection(subcollectionNames.SUBSCRIPTIONS)
        .doc(activityId),
      activityNewData,
      { merge: true },
    );
  }
};

/**
 * Handle templates with role parameter
 * @param locals
 * @return {Promise<null|void>}
 */
const handleRoleDoc = async locals => {
  const batch = db.batch();
  const { id: activityId } = locals.change.after;
  const activityNewData = locals.change.after.data();
  const { status, template } = activityNewData;
  if (status === 'CONFIRMED') {
    await handleRoleDocConfirmed(activityNewData, batch, template, activityId);
  }
  if (status === 'CANCELLED') {
    await handleRoleDocCancelled(activityNewData, batch, template, activityId);
  }
  try {
    await batch.commit();
  } catch (error) {
    console.error(error);
  }
  return null;
};

const templateHandler = async locals => {
  const { template } = locals.change.after.data();
  const action = locals.addendumDocData ? locals.addendumDocData.action : null;

  if (template === 'check-in' || action === httpsActions.checkIn) {
    await reimburseDailyAllowance(locals);
    await reimburseKmAllowance(locals);

    return handleWorkday(locals);
  }

  if (template === 'office') {
    await require('./template-handlers/office')(locals);
  }

  if (template === 'recipient') {
    await handleRecipient(locals);
  }

  if (template === 'subscription') {
    try {
      await handleSubscription(locals);
    } catch (error) {
      console.error(error);
    }
  }

  if (template === 'admin') {
    await require('./template-handlers/admin')(locals);
  }

  const activityReportName = await getActivityReportName({
    template,
    report: locals.change.after.get('report'),
  });

  if (activityReportName === 'role') {
    await handleRoleDoc(locals);
  }
  if (activityReportName === 'attendance') {
    await attendanceHandler(locals);
  }

  if (activityReportName === 'reimbursement') {
    await handleReimbursement(locals);
  }

  await handleConfig(locals);
  await handleTypeActivity(locals);
  await handleMetaUpdate(locals);

  await setLocationsReadEvent(locals);

  // await mapActivityToUserUpdates(locals.change.after, null);

  // handle growthfileMs integration parallely
  await growthFileMsIntegration(locals.change).catch(console.error);

  return;
};

/**
 * @param {{ after: any; before?: { data: () => any; }; }} change
 */
const handleProfile = async change => {
  const batch = db.batch();
  const {
    after: { id: activityId },
  } = change;
  const newPhoneNumbersSet = new Set();
  const authFetchPromises = [];
  const { template, addendumDocRef } = change.after.data();
  const locals = {
    change,
    /** Used while creating comment */
    activityNew: change.after,
    addendumDoc: null,
    assigneesMap: new Map(),
    assigneePhoneNumbersArray: [],
    addendumCreator: {},
    addendumCreatorInAssignees: false,
  };

  const promises = [
    rootCollections.activities
      .doc(activityId)
      .collection(subcollectionNames.ASSIGNEES)
      .get(),
  ];

  /** Could be `null` when we update the activity without user intervention */
  if (addendumDocRef) {
    promises.push(addendumDocRef.get());
  }

  const [assigneesSnapShot, addendumDoc] = await Promise.all(promises);

  locals.addendumDoc = addendumDoc || null;

  assigneesSnapShot.forEach(doc => {
    const { id: phoneNumber } = doc;
    const { addToInclude } = doc.data();

    // @ts-ignore
    if (addendumDoc && phoneNumber === addendumDoc.get('user')) {
      locals.addendumCreatorInAssignees = true;
    }

    authFetchPromises.push(getAuth(phoneNumber));

    /** Storing phoneNumber in the object because we are
     * storing assigneesMap in addendum doc
     **/
    locals.assigneesMap.set(phoneNumber, {
      phoneNumber,
      addToInclude: addToInclude || false,
    });

    locals.assigneePhoneNumbersArray.push(phoneNumber);
  });

  if (addendumDoc && !locals.addendumCreatorInAssignees) {
    // @ts-ignore
    authFetchPromises.push(getAuth(addendumDoc.get('user')));
  }

  locals.customerObject = await getLocationObject({
    name: change.after.get('attachment.Location.value'),
    officeId: change.after.get('officeId'),
  });

  const userRecords = await Promise.all(authFetchPromises);

  userRecords.forEach(userRecord => {
    const { phoneNumber, uid } = userRecord;

    if (
      addendumDoc &&
      !locals.addendumCreatorInAssignees &&
      phoneNumber === addendumDoc.get('user')
    ) {
      locals.addendumCreator.displayName = userRecord.displayName;

      /**
       * Since addendum creator was not in the assignees list,
       * returning from the iteration since we don't want to
       * add them to the activity unnecessarily.
       */
      return;
    }

    locals.assigneesMap.set(
      phoneNumber,
      Object.assign({}, locals.assigneesMap.get(phoneNumber), {
        uid: uid || '',
        displayName: userRecord.displayName || '',
        photoURL: userRecord.photoURL || '',
        customClaims: userRecord.customClaims || {},
      }),
    );

    /** New user introduced to the system. Saving their phone number. */
    if (!uid) {
      newPhoneNumbersSet.add(phoneNumber);
    }
  });

  const profileActivityObject = getProfileActivityObject({
    customerObject: locals.customerObject,
    activityDoc: change.after,
    assigneePhoneNumbersArray: locals.assigneePhoneNumbersArray,
    assigneesMap: locals.assigneesMap,
  });

  userRecords.forEach(({ uid, phoneNumber }) => {
    // Check-ins are sent to users via `Updates/{uid}/Addendum/` collection
    if (template !== 'check-in') {
      // in profile
      if (!phoneNumber) {return;}
      batch.set(
        rootCollections.profiles
          .doc(phoneNumber)
          .collection(subcollectionNames.ACTIVITIES)
          .doc(activityId),
        profileActivityObject,
        { merge: true },
      );
    }

    if (uid) {
      // in updates only if auth exists
      batch.set(
        rootCollections.updates
          .doc(uid)
          .collection(subcollectionNames.ADDENDUM)
          .doc(),
        Object.assign({}, profileActivityObject, {
          _type: addendumTypes.ACTIVITY,
        }),
      );
    }
  });

  console.log({
    template,
    activityId,
    action: locals.addendumDoc
      ? locals.addendumDoc.get('action')
      : 'manual update',
  });

  const copyToRef = getCopyPath({
    template,
    activityId,
    officeId: change.after.get('officeId'),
  });

  batch.set(
    copyToRef,
    getSubcollectionActivityObject({
      activityObject: change.after,
      customerObject: locals.customerObject,
    }),
    { merge: true },
  );

  await batch.commit();

  await createNewProfiles({
    newPhoneNumbersSet,
    smsContext: {
      activityName: change.after.get('activityName'),
      office: change.after.get('office'),
      creator:
        change.after.get('creator.phoneNumber') || change.after.get('creator'),
    },
  });

  return locals;
};

/**
 * @param {{ after: { data: () => { (): any; new (): any; template: string; }; }; before: { data: () => any; }; }} change
 */
const activityOnWrite = async change => {
  /** Activity was deleted. For debugging only. */
  if (!change.after.data()) {
    return;
  }

  /**
   * The sequence of handleAddendum and handleProfile matters for
   * correct execution flow. All other functions can be called in
   * any order (mostly).
   */
  const locals = await handleProfile(change);

  /*
    @see BugFix
    This function checks for updated phone numbers except employee's phone number and then
    triggers the same change in subscription activities of the employee
  */
  if (change.before.data() && change.after.data().template === 'employee') {
    await handleSupervisorUpdate(locals);
  }

  await handleAddendum(locals);

  await templateHandler(locals);

  return handleComments(locals.addendumDoc, locals);
};

module.exports = (change, context) => {
  try {
    return activityOnWrite(change);
  } catch (error) {
    console.error({
      error,
      context,
      activityId: change.after.id,
    });
  }
};
