(function () {
  if (!sessionStorage.getItem('prefill-form')) {
    return;
  }

  const {
    officeName,
    firstContact,
    firstContactDisplayName,
    firstContactEmail,
    secondContact,
    secondContactDisplayName,
    secondContactEmail,
  } = JSON.parse(sessionStorage.getItem('prefill-form'));

  if (officeName) {
    form.elements.namedItem('office-name').value = officeName;
  }

  if (firstContact) {
    form.elements.namedItem('user-phone-number').value = firstContact;
  }

  if (firstContactDisplayName) {
    form.elements.namedItem('user-name').value = firstContactDisplayName;
  }

  if (firstContactEmailElement) {
    form.elements.namedItem('user-email').value = firstContactEmailElement
  }

  if (secondContactElement) {
    form.elements.namedItem('second-contact-phone-number').value = secondContactElement;
  }

  if (secondContactDisplayNameElement) {
    form.elements.namedItem('second-contact-name').value = secondContactDisplayNameElement;
  }

  if (secondContactEmailElement) {
    form.elements.namedItem('second-contact-email').value = secondContactEmailElement;
  }
})();


function validateForm() {
  const form = document.forms[0];
  const officeNameElement = form.elements.namedItem('office-name');
  const firstContactElement = form.elements.namedItem('user-phone-number');
  const firstContactDisplayNameElement = form.elements.namedItem('user-name');
  const firstContactEmailElement = form.elements.namedItem('user-email');
  const secondContactElement = form.elements.namedItem('second-contact-phone-number');
  const secondContactDisplayNameElement = form.elements.namedItem('second-contact-name');
  const secondContactEmailElement = form.elements.namedItem('second-contact-email');

  let valid = true;

  if (!isNonEmptyString(officeNameElement.value)) {
    const element = getWarningNode('Office Name');

    insertAfterNode(officeNameElement, element);
  }

  if (!isNonEmptyString(firstContactElement.value)) {
    const element = getWarningNode('Your Phone Number');

    insertAfterNode(firstContactElement, element);
  }

  if (!isNonEmptyString(firstContactDisplayNameElement.value)) {
    const element = getWarningNode('Your name');

    insertAfterNode(firstContactDisplayNameElement, element);
  }

  if (!isNonEmptyString(firstContactEmailElement.value)) {
    const element = getWarningNode('Your Email');

    insertAfterNode(firstContactEmailElement, element);
  }

  if (!isNonEmptyString(secondContactElement.value)) {
    const element = getWarningNode('Second Contact');

    insertAfterNode(secondContactElement, element);
  }

  if (!isNonEmptyString(secondContactDisplayNameElement.value)) {
    const element = getWarningNode('Second Contact Name');

    insertAfterNode(secondContactDisplayNameElement, element);
  }

  if (!isNonEmptyString(secondContactEmailElement.value)) {
    const element = getWarningNode('Second Contact Email');

    insertAfterNode(secondContactEmailElement, element);
  }

  return {
    valid,
    values: {
      officeName: officeNameElement.value,
      firstContactPhoneNumber: firstContactElement.value,
      secondContactElementPhoneNumber: secondContactElement.value,
      firstContactDisplayName: firstContactDisplayNameElement.value,
      secondContactDisplayName: secondContactDisplayNameElement.value,
      firstContactEmail: firstContactEmailElement.value,
      secondContactEmail: secondContactElement.value,
    },
  }
}

function sendOfficeCreationRequest(values) {
  // console.log('creating office');
  if (!startPosition) {
    return askLocationPermission({}, sendOfficeCreationRequest);
  }

  const spinner = getSpinnerElement();
  document.forms[0].innerText = '';
  document.forms[0].style.display = 'flex';
  document.forms[0].style.justifyContent = 'center';

  spinner.id = 'join-fetch-spinner';

  document.forms[0].appendChild(spinner);

  const requestUrl = '';
  const requestBody = {
    timestamp: Date.now(),
    office: '',
    template: 'office',
    geopoint: {
      // latitude: 
      // longitude:
    },
    data: [{
      Name: '',
      Description: '',
      'Youtube ID': '',
      'GST Number': '',
      'First Contact': '',
      'Second Contact': '',
      Timezone: moment.tz.guess(),
      'Head Office': '',
      'Date of Establishment': '',
      'Trial Period': '',
    }],
  }

  const idToken = getParsedCookies().__session;

  // const requestUrl = 'https://api2.growthfile.com/api/activities/create';
  const requestUrl = 'https://us-central1-growthfilev2-0.cloudfunctions.net/api/admin/bulk';

  return fetch(requestUrl, {
    mode: 'cors',
    method: 'POST',
    body: JSON.stringify(requestBody),
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${idToken}`,
    },
  })
    .then((result) => result.json())
    .then((response) => {
      console.log('Response', response);

      document
        .getElementById('join-fetch-spinner')
        .style.display = 'none';

      const span = document.createElement('span');

      let spanText = 'Office Created Successfully';

      if (!response.success) {
        spanText = response.message;
        span.classList.add('warning-label');
      } else {
        span.classList.add('success-label');
      }

      span.innerHTML = spanText;
      document.forms[0].appendChild(span);

      // redirect to the home page
      return;
    })
    .catch(console.error);
};

function startOfficeCreationFlow(event) {
  event.preventDefault();

  const oldWarningLabels = document.querySelectorAll('p .warning-label');

  Array
    .from(oldWarningLabels)
    .forEach((element) => element.style.display = 'none');

  const result = validateForm();

  if (!result.valid) return;

  uiConfig
    .signInOptions[0]
    .defaultNationalNumber = document
      .forms[0]
      .elements
      .namedItem('user-phone-number')
      .value;

  /** Not logged-in */
  if (!firebase.auth().currentUser) {
    const modal = showLoginBox('90%', 'fb-auth-modal');

    modal.how();

    return;
  }

  return sendOfficeCreationRequest(result.values);
}

document.addEventListener('onbeforeunload', function () {
  console.log('saving form data to sessionstorage');

  const form = document.forms[0];
  const officeNameElement = form.elements.namedItem('office-name');
  const firstContactElement = form.elements.namedItem('user-phone-number');
  const firstContactDisplayNameElement = form.elements.namedItem('user-name');
  const firstContactEmailElement = form.elements.namedItem('user-email');
  const secondContactElement = form.elements.namedItem('second-contact-phone-number');
  const secondContactDisplayNameElement = form.elements.namedItem('second-contact-name');
  const secondContactEmailElement = form.elements.namedItem('second-contact-email');

  sessionStorage
    .setItem('prefill-form', JSON.stringify({
      officeName: officeNameElement.value,
      firstContact: firstContactElement.value,
      firstContactDisplayName: firstContactDisplayNameElement.value,
      firstContactEmail: firstContactEmailElement.value,
      secondContact: secondContactElement.value,
      secondContactDisplayName: secondContactDisplayNameElement.value,
      secondContactEmail: secondContactEmailElement.value,
    }));
});
