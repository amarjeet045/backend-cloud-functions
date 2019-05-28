'use strict';

function handleRecaptcha() {
  return new firebase.auth.RecaptchaVerifier('recaptcha-container', {
    'size': 'normal'
  });
}


function isValidPhoneNumber(phoneNumber = '') {
  const pattern = /^\+[0-9\s\-\(\)]+$/;

  return phoneNumber.search(pattern) !== -1;
}

function getParsedCookies() {
  const cookieObject = {};

  document
    .cookie
    .split(';')
    .forEach((cookie) => {
      const parts = cookie.split('=');

      cookieObject[parts.shift().trim()] = decodeURI(parts.join('='));
    });

  return cookieObject;

};

function isNonEmptyString(string) {
  return typeof string === 'string' && string.trim() !== '';
}

function insertAfterNode(currentNode, nodeToInsert) {
  currentNode.parentNode.insertBefore(nodeToInsert, currentNode.nextSibling);
}

function logoutUser(event) {
  event.preventDefault();

  /** User isn't logged in */
  if (!firebase.auth().currentUser) return;

  console.log('logging out user...');

  document.cookie = `__session=`;

  return firebase
    .auth()
    .signOut()
    .then(function () {
      window.location.reload();

      return;
    })
    .catch(console.error);
};

function getWarningNode(textContent) {
  // valid = false;

  const warningNode = document.createElement('span');
  warningNode.classList.add('warning-label');
  warningNode.textContent = textContent;

  return warningNode;
}

function getQueryString(field, url) {
  const href = url ? url : window.location.href;
  const reg = new RegExp('[?&]' + field + '=([^&#]*)', 'i');
  const string = reg.exec(href);

  return string ? string[1] : null;
}

function getMobileOperatingSystem() {
  var userAgent = navigator.userAgent || navigator.vendor || window.opera;

  // Windows Phone must come first because its UA also contains "Android"
  if (/windows phone/i.test(userAgent)) {
    return 'Windows Phone';
  }

  if (/android/i.test(userAgent)) {
    return 'Android';
  }

  // iOS detection from: http://stackoverflow.com/a/9039885/177710
  if (/iPad|iPhone|iPod/.test(userAgent) && !window.MSStream) {
    return 'iOS';
  }

  return 'unknown';
};

function isValidEmail(emailString) {
  return /^([A-Za-z0-9_\-\.])+\@([A-Za-z0-9_\-\.])+\.([A-Za-z]{2,4})$/
    .test(emailString);
}

function getSpinnerElement(id) {
  const elem = document.createElement('div');
  elem.className = 'spinner';
  elem.style.position = 'relative';
  elem.style.height = '40px';
  elem.style.width = '40px';

  if (id) {
    elem.id = id;
  }
  return {
    center: function () {
      elem.classList.add('spinner-center')
      return elem;
    },
    default: function () {
      return elem;
    }
  }

}

/** Create Modal box */
function createModal(actionContent) {
  if (document.getElementById('modal')) {
    // document.getElementById('modal').remove();
    setContentInModal(actionContent, document.querySelector('#modal .action-container'))
    return;
  };

  const div = document.createElement('div');
  div.className = 'modal';
  div.id = 'modal'


  const content = document.createElement('div')
  content.className = 'modal-content';

  const close = document.createElement('span')
  close.className = 'close fa fa-window-close'
  close.onclick = function () {
    div.remove();
  }
  content.appendChild(close)

  const actionContainer = document.createElement('div')
  actionContainer.className = 'action-container mt-10';
  const actionNotification = document.createElement('p');
  actionNotification.id = 'action-label'
  content.appendChild(actionNotification)
  setContentInModal(actionContent, actionContainer)
  content.appendChild(actionContainer);
  div.appendChild(content)
  return div;
}

function isDomElementString(el) {
  return typeof el == 'string';
}

function setContentInModal(el, parent) {
  console.log(el)
  console.log(parent);

  if (isDomElementString(el)) {
    parent.innerHTML = el;
  } else {
    parent.appendChild(el);
  }
}

function setMessage(message) {
  const messageNode = document.getElementById('message');
  messageNode.innerText = message;
  messageNode.classList.remove('hidden');
}


function getLocation() {
  return new Promise((resolve, reject) => {
    if (!navigator.geolocation) return reject('Geolocation is Not Supported')
    navigator
      .geolocation
      .getCurrentPosition(function (position) {
        return resolve({
          latitude: position.coords.latitude,
          longitude: position.coords.longitude,
          accuracy: position.coords.accuracy
        })
      }, function (error) {
        let message;
        switch (error.code) {
          case 1:
            message = 'Please Enable Location';
            break;
          default:
            message = error.message;
        }
        return reject(message)
      });
  })
}


function sendApiRequest(apiUrl, requestBody, method) {
  const init = {
    method,
    mode: 'cors',
    cache: 'no-cache',
    headers: {
      'Content-Type': 'application/json',
    },
  };

  if (requestBody && init.method !== 'GET') {
    init.body = JSON.stringify(requestBody);
  }

  return firebase
    .auth()
    .currentUser
    .getIdToken(false)
    .then(function (idToken) {
      init.headers['Authorization'] = `Bearer ${idToken}`;

      return fetch(apiUrl, init);
    })
    .then(function (result) {
      return result;
    })
    .catch(console.error);
}


document.addEventListener('click', (event) => {
  if (event.target === document.getElementById('form-submit-button')) {
    return void startOfficeCreationFlow(event)
  }

  if (event.target === document.getElementById('load-map-button')) {
    return getLocation().then(initMap).catch(function (message) {
      if (document.getElementById('map')) {
        document.getElementById('map').innerHTML = `<p style='text-align:center;margin-top:20px;' class='warning-label'>${message}</p>`
      }
    })
  }

  if (event.target === document.getElementById('enquiry-submit-button')) {
    return void startEnquiryCreationFlow(event);
  }

  // TODO: Refactor this name. Not very unique and might cause conflicts.
  if (Array.from(document.querySelectorAll('.list-item')).includes(event.target)) {
    return void updateMapPointer(event);
  }

  if (event.target === document.querySelector('#header-hamburger-icon')) {
    document.querySelector('aside').classList.toggle('hidden');
  }


  if (event.target === document.getElementById('menu-logout-link')) {
    return void logoutUser(event);
  }
});

firebase
  .auth()
  .onAuthStateChanged(function (user) {
    if (user) return;

    document.cookie = `__session=`;
    console.log('no session cookie');
  });

function setGlobals() {
  console.log('fetching config');

  return fetch('/config')
    .then(function (response) { return response.json() })
    .then(function (result) {
      Object
        .keys(result)
        .forEach(function (key) {
          window[key] = result[key];
        });

      console.log('config set:', result);
    })
    .catch(console.error);
}

function checkDnt() {
  const dntEnabled = navigator.doNotTrack === 1;

  console.log({ dntEnabled });
}


window
  .addEventListener('load', function () {
    setGlobals();
    checkDnt();

    firebase
      .auth()
      .addAuthTokenListener(function (idToken) {
        if (!idToken) return;

        document.cookie = `__session=${idToken};max-age=${idToken ? 3600 : 0};`;

        console.log('new cookie set', idToken);
      });
  });
