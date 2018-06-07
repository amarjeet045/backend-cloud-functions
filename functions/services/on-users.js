const {
  parse,
} = require('url');

const {
  sendResponse,
} = require('../admin/utils');

const {
  code,
} = require('../admin/responses');

const onRead = require('../services/users/on-read');
const onUpdate = require('../services/users/on-update');


/**
 * Handles the requests made to /users resource.
 *
 * @param {Object} conn Contains Express' Request and Respone objects.
 */
const app = (conn) => {
  const action = parse(conn.req.url).path.split('/')[3];

  if (action.startsWith('read')) {
    if (conn.req.method !== 'GET') {
      sendResponse(
        conn,
        code.methodNotAllowed,
        `${conn.req.method} is not allowed for the /read endpoint.`
      );
      return;
    }

    onRead(conn);
    return;
  }

  if (action === 'update') {
    if (conn.req.method !== 'PATCH') {
      sendResponse(
        conn,
        code.methodNotAllowed,
        `${conn.req.method} is not allowed for the /update endpoint.`
      );
      return;
    }

    onUpdate(conn);
    return;
  }

  sendResponse(conn, code.badRequest, 'The request path is not valid.');
};


module.exports = app;
