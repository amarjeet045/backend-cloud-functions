'use strict';

const {
  sendResponse,
  hasSupportClaims,
  disableAccount,
} = require('../../admin/utils');
const {
  code,
} = require('../../admin/responses');
const {
  isValidRequestBody,
} = require('../activity/helper');
const {
  httpsActions,
} = require('../../admin/constants');


module.exports = (conn) => {
  /**
   * Unlike the client-side APIs, this one simply replaces
   * or creates the documents.
   */
  if (conn.req.method !== 'PUT') {
    sendResponse(
      conn,
      code.methodNotAllowed,
      `${conn.req.method} is not allowed. Use 'PUT'`
    );

    return;
  }

  const result =
    isValidRequestBody(conn.req.body, httpsActions.create);

  if (!result.isValid) {
    sendResponse(conn, code.badRequest, result.message);

    return;
  }

  /** Only `support` can create an office */
  if (conn.req.body.template === 'office'
    && !hasSupportClaims(conn.requester.customClaims)) {
    // disableAccount(conn, 'You do not have the permission to access this resource');

    sendResponse(conn, 'You do not have the permission to access this resource');

    return;
  }

  if (!new Set()
    .add('office')
    .add('product')
    .add('recipient')
    .add('department')
    .add('leave-type')
    .add('expense-type')
    .add('supplier-type')
    .add('customer-type')
    .has(conn.req.body.template)) {
    sendResponse(
      conn,
      code.forbidden,
      `Template: '${conn.req.body.template}' is not supported`
    );

    return;
  }

  if (conn.req.body.hasOwnProperty('activityId')) {
    require('./update')(conn);

    return;
  }

  require('./create')(conn);
};