'use strict';

const admin = require('firebase-admin');

const {
  rootCollections,
} = require('../../../admin/admin');



module.exports = (snapShot, context) =>
  rootCollections
    .updates
    .doc(context.params.uid)
    .get()
    .then((updatesDoc) => {
      const registrationToken = updatesDoc.get('registrationToken');

      if (!registrationToken) {
        console.log('NO REGESTRATION TOKEN FOUND. EXITING...');

        return Promise.resolve();
      }

      const payload = {
        data: {
          // Ask the client to send a request to the /read endpoint
          read: '1',
        },
        notification: {
          body: snapShot.get('comment'),
          tile: `Growthfile`,
        },
      };
      const ONE_MINUTE = 60;
      const options = {
        priority: 'high',
        timeToLive: ONE_MINUTE,
      };

      console.log(`Notification sent to `
        + `phoneNumber=${updatesDoc.get('phoneNumber')},`
        + ` 'uid=${context.params.uid}'`, { payload });

      return admin
        .messaging()
        .sendToDevice(registrationToken, payload, options);
    })
    .then(JSON.stringify)
    .catch(console.error);
