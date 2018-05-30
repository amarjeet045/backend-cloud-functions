const {
  users,
} = require('../../admin/admin');

const {
  sendResponse,
  handleError,
} = require('../../admin/utils');

const {
  isValidPhoneNumber,
} = require('../../firestore/activity/helper');

const {
  code,
} = require('../../admin/responses');

const {
  setCustomUserClaims,
  getUserByPhoneNumber,
} = users;


const setClaims = (conn, userRecord) => {
  const permissions = {};

  if (typeof conn.req.body.permissions.support === 'boolean') {
    permissions.support = conn.req.body.permissions.support;
  }

  if (typeof conn.req.body.permissions.manageTemplates === 'boolean') {
    permissions.manageTemplates = conn.req.body.permissions.manageTemplates;
  }

  /** Both keys are not present in the request body. Skip giving
   * the permissions.
   * */
  if (Object.keys(permissions).length < 1) {
    sendResponse(
      conn,
      code.badRequest,
      'The "permissions" object in the request body is invalid.'
    );
    return;
  }

  setCustomUserClaims(userRecord.uid, permissions).then(() => {
    sendResponse(
      conn,
      code.ok,
      `Updated permissions for ${conn.req.body.phoneNumber} successfully.`
    );
    return;
  }).catch((error) => handleError(conn, error));
};


const getUserRecordFromPhoneNumber = (conn) => {
  getUserByPhoneNumber(conn.req.body.phoneNumber).then((userRecord) => {
    if (!userRecord[conn.req.body.phoneNumber].uid) {
      sendResponse(
        conn,
        code.badRequest,
        `No user with phone number ${conn.req.body.phoneNumber} exists.`);
      return;
    }

    setClaims(conn, {
      uid: userRecord[conn.req.body.phoneNumber].uid,
    });
    return;
  }).catch((error) => handleError(conn, error));
};


const validateRequestBody = (conn) => {
  if (!isValidPhoneNumber(conn.req.body.phoneNumber)) {
    sendResponse(
      conn,
      code.badRequest,
      `${conn.req.body.phoneNumber} is not a valid phone number.`
    );
    return;
  }

  if (conn.requester.phoneNumber === conn.req.body.phoneNumber) {
    sendResponse(
      conn,
      code.forbidden,
      'You cannot change your own permissions.'
    );
    return;
  }

  if (!conn.req.body.permissions) {
    sendResponse(
      conn,
      code.badRequest,
      'The "permisssions" object is missing from the request body.'
    );
    return;
  }

  /** The only object we need here is {...}. Anything else like an `Array`
   * is should is not allowed.
   */
  if (Object.prototype.toString
    .call(conn.req.body.permissions) !== '[object Object]') {
    sendResponse(
      conn,
      code.badRequest,
      'The "permissions" object in request body is invalid.'
    );
    return;
  }

  getUserRecordFromPhoneNumber(conn);
};


const app = (conn) => {
  /** The superUser field can be `undefined', true` or `false`. So, we are
    * explictly checking its value against a boolean. Only a `truthy`/`falsey`
    * check will probably not work here.
    */
  if (conn.requester.customClaims.superUser !== true) {
    sendResponse(
      conn,
      code.forbidden,
      'You are unauthorized from changing permissions.'
    );
    return;
  }

  validateRequestBody(conn);
};


module.exports = app;