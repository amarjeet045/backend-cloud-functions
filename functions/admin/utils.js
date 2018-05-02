const now = (conn) => {
  if (conn.req.method === 'GET') {
    sendResponse(conn, 200, (new Date()).toUTCString());
  } else {
    sendResponse(conn, 405, 'METHOD NOT ALLOWED');
  }
};

const sendResponse = (conn, statusCode, statusMessage) => {
  conn.headers['Content-Type'] = 'application/json';
  conn.res.writeHead(statusCode, conn.headers);

  conn.res.end(JSON.stringify({
    message: statusMessage,
  }));
};

const handleError = (conn, error) => {
  console.log(error);
  sendResponse(conn, 500, 'INTERNAL SERVER ERROR');
};


module.exports = {
  sendResponse,
  handleError,
  now,
};