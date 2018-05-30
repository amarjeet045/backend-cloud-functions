const {
  getGeopointObject,
  rootCollections,
} = require('../../admin/admin');

const {
  activityTemplates,
} = rootCollections;

const {
  code,
} = require('../../admin/responses');

const {
  isValidLocation,
  isValidDate,
} = require('../activity/helper');

const {
  handleError,
  sendResponse,
} = require('../../admin/utils');


const validateSchedule = (schedule) => {
  if (typeof schedule.name !== 'string') return false;
  if (!isValidDate(schedule.startTime)) return false;
  if (!isValidDate(schedule.endTime)) return false;
  if (schedule.endTime < schedule.startTime) return false;

  return true;
};

const validateVenue = (venue) => {
  if (typeof venue.venueDescriptor !== 'string') return false;
  if (typeof venue.address !== 'string') return false;
  if (typeof venue.location !== 'string') return false;
  if (!isValidLocation(venue.geopoint)) return false;

  return true;
};


const createTemplate = (conn, data) => {
  activityTemplates.doc().set(data).then(() => {
    sendResponse(
      conn,
      code.created,
      'The template was created successfully.'
    );
    return;
  }).catch((error) => handleError(conn, error));
};


const createTemplateObject = (conn) => {
  const data = {
    comment: conn.req.body.comment,
    name: conn.req.body.name,
    defaultTitle: conn.req.body.defaultTitle,
    venue: {
      address: conn.req.body.venue.address,
      geopoint: getGeopointObject(conn.req.body.venue.geopoint),
      location: conn.req.body.venue.location,
      venueDescriptor: conn.req.body.venue.venueDescriptor,
    },
    schedule: {
      name: conn.req.body.schedule.name,
      startTime: new Date(conn.req.body.schedule.startTime),
      endTime: new Date(conn.req.body.schedule.endTime),
    },
    attachment: conn.req.body.attachment,
  };

  createTemplate(conn, data);
};


const getTemplateByName = (conn) => {
  activityTemplates.where('name', '==', conn.req.body.name).limit(1)
    .get().then((snapShot) => {
      if (!snapShot.empty) {
        sendResponse(
          conn,
          code.conflict,
          `A template with the name: ${conn.req.body.name} already exists.`
        );
        return;
      }

      createTemplateObject(conn);
      return;
    }).catch((error) => handleError(conn, error));
};


const validateRequestBody = (conn) => {
  /** true for empty string. */
  const re = /^$|\s+/;

  if (!re.test(conn.req.body.name)) {
    sendResponse(conn, code.badRequest, 'The "name" is invalid/missing');
    return;
  }

  if (!re.test(conn.req.body.defaultTitle)) {
    sendResponse(conn, code.badRequest, 'The "description" is invalid/missing');
    return;
  }

  if (!re.test(conn.req.body.comment)) {
    sendResponse(conn, code.badRequest, 'The "comment" is invalid/missing');
    return;
  }

  if (Object.prototype.toString
    .call(conn.req.body.attachment) !== '[object Object]') {
    sendResponse(
      conn,
      code.badRequest,
      'The "template" in invalid/missing'
    );
    return;
  }

  if (Object.prototype.toString
    .call(conn.req.body.schedule) !== '[object Object]') {
    sendResponse(
      conn,
      code.badRequest,
      'The "schedule" is missing/invalid'
    );
    return;
  }

  if (Object.prototype.toString
    .call(conn.req.body.venue) !== '[object Object]') {
    sendResponse(
      conn,
      code.badRequest,
      'The "venue" is invalid/missing'
    );
    return;
  }

  if (!validateSchedule(conn.req.body.schedule)) {
    sendResponse(
      conn,
      code.badRequest,
      'The "schedule" object is invalid'
    );
    return;
  }

  if (!validateVenue(conn.req.body.venue)) {
    sendResponse(
      conn,
      code.badRequest,
      'The "venue" object is invalid'
    );
    return;
  }

  getTemplateByName(conn);
};


const app = (conn) => {
  validateRequestBody(conn);
};


module.exports = app;
