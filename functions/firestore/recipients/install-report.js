'use strict';

const {
  sendGridTemplateIds,
} = require('../../admin/constants');
const {
  rootCollections,
} = require('../../admin/admin');

const {
  getYesterdaysDateString,
} = require('./report-utils');

/**
 * Returns yesterday's Day start timestamp.
 * @returns {Object} JS date object of the previous day starting timestamp.
 */
const getYesterdaysStartTime = () => {
  const today = new Date();
  today.setHours(0, 0, 0);

  return new Date(today.setDate(today.getDate() - 1));
};


const getName = (employeesData, phoneNumber) => {
  if (!employeesData[phoneNumber]) return '';

  return employeesData[phoneNumber].Name;
};


module.exports = (locals) => {
  const {
    office,
    officeId,
  } = locals.change.after.data();

  const yesterdaysDateString = getYesterdaysDateString();

  locals.messageObject.templateId = sendGridTemplateIds.installs;
  locals.csvString =
    ` Employee Name,`
    + ` Employee Contact,`
    + ` Employee Code,`
    + ` Department,`
    + ` Installed On,`
    + ` Number Of Installs,`
    + ` First Supervisor's Name,`
    + ` Contact Number,`
    + ` Second Supervisor's Name,`
    + ` Contact Number\n`;

  locals.messageObject['dynamic_template_data'] = {
    office,
    date: yesterdaysDateString,
    subject: `Install Report_${office}_${yesterdaysDateString}`,
  };

  locals.multipleInstallsMap = new Map();

  Promise
    .all([
      rootCollections
        .offices
        .doc(officeId)
        .get(),
      rootCollections
        .inits
        .where('office', '==', office)
        .where('report', '==', 'install')
        .where('date', '==', yesterdaysDateString)
        .get(),
    ])
    .then((result) => {
      const [
        officeDoc,
        installDocs,
      ] = result;

      if (installDocs.empty) {
        /** No report to be sent since no one installed yesterday. */
        return Promise.resolve();
      }

      let totalInstalls = 0;

      // Collecting the list of people who have multiple installs for yesterday.
      const yesterdaysStartTime = getYesterdaysStartTime();

      let header = 'Install Date and Time\n\n';

      installDocs.forEach((doc) => {
        const {
          phoneNumber,
          installs,
        } = doc.data();

        const employeeData =
          officeDoc
            .get('employeesData')[phoneNumber];

        installs
          .forEach((timestampString) => header += `${timestampString}\n`);

        installs.forEach((timestampString) => {
          totalInstalls++;

          const installTime =
            new Date(timestampString).getTime();

          if (installTime > yesterdaysStartTime) return;

          locals.multipleInstallsMap.set(phoneNumber, header);
        });

        const firstSupervisorPhoneNumber =
          employeeData['First Supervisor'];
        const secondSupervisorPhoneNumber =
          employeeData['Second Supervisor'];
        const firstSupervisorsName =
          getName(officeDoc.get('employeesData'), firstSupervisorPhoneNumber);
        const secondSupervisorsName =
          getName(officeDoc.get('employeesData'), secondSupervisorPhoneNumber);

        locals.csvString +=
          ` ${employeeData.Name},`
          + ` ${phoneNumber},`
          + ` ${employeeData['Employee Code']},`
          + ` ${employeeData.Department},`
          + ` ${installs[installs.length - 1]},`
          + ` ${installs.length},`
          + ` ${firstSupervisorsName},`
          + ` ${firstSupervisorPhoneNumber},`
          + ` ${secondSupervisorsName},`
          + ` ${secondSupervisorPhoneNumber}`
          + `\n`;
      });

      locals
        .messageObject['dynamic_template_data']
        .totalInstalls = totalInstalls;

      locals.messageObject['dynamic_template_data']
        .extraInstalls = totalInstalls - installDocs.size;

      locals.messageObject.attachments.push({
        content: new Buffer(locals.csvString).toString('base64'),
        fileName: `${office} Install Report_${yesterdaysDateString}.csv`,
        type: 'text/csv',
        disposition: 'attachment',
      });

      locals
        .multipleInstallsMap
        .forEach((timestampsString, phoneNumber) => {
          locals.messageObject.attachments.push({
            content: new Buffer(timestampsString).toString('base64'),
            fileName: `${phoneNumber}.txt`,
            type: 'text/plain',
            disposition: 'attachment',
          });
        });

      console.log({
        locals,
      });

      return locals
        .sgMail
        .sendMultiple(locals.messageObject);
    })
    .catch(console.error);
};
