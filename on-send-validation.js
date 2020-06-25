/* global Office */

const webConferencingDomains = new Array(
    'zoom\\.us',
    'zoom\\.com',
    'webex\\.com',
    'bluejeans\\.com',
    'appointlet\\.com',
    'join\\.me',
    'chime\\.aws',
    'clearslide\\.com',
    'linkedinslides\\.com',
    'gotomeeting\\.com',
    'gotomeet\\.me',
    'meet\\.google\\.com',
    'uberconference\\.com',
    'lync\\.com',
    'meetings\\.ringcentral\\.com',
    'teams\\.microsoft\\.com');

const webConferencingRegexCheck = new RegExp('\\b(' + webConferencingDomains.join('|') + ')\\b');
const linkedinDomains = new Array('linkedin\\.com$', 'linkedin\\.biz$', 'glintinc\\.com$');
const linkedinDomainsRegexCheck = new RegExp(linkedinDomains.join("|"));

const GONG_COORDINATOR = { NAME: 'Gong Coordinator', EMAIL: 'coordinator@gong.io' };

let mailboxItem;
let containsExternalAttendee;

Office.initialize = function () {
    mailboxItem = Office.context.mailbox.item;
};

function addGongCoordinator(event) {
    mailboxItem.optionalAttendees.addAsync([{
        displayName: GONG_COORDINATOR.NAME,
        emailAddress: GONG_COORDINATOR.EMAIL,
    }], { asyncContext: event });
    event.completed({ allowEvent: true });
}

function confirmationAndAddCoordinator(event) {
    const addGongAutomatically = localStorage.getItem('linkedin-add-in_always_invite_gong');
    if (addGongAutomatically != null && addGongAutomatically === 'true') {
        addGongCoordinator(event);
    } else {
        Office.context.ui.displayDialogAsync('https://gong-for-linkedin.s3.amazonaws.com/validation-dialog.html', {
            height: 70,
            width: 30,
            displayInIframe: true,
        }, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
            } else {
                const dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (messageEvent) {
                    dialog.close();
                    if (messageEvent.message === 'yes') {
                        addGongCoordinator(event);
                    } else {
                        event.completed({ allowEvent: true });
                    }
                });
                dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
                    dialog.close();
                    event.completed({ allowEvent: true });
                });
            }
        });
    }
}

function check4Body(asyncResult) {
    const body = asyncResult.value;
    const lowerCaseBody = body.toLowerCase();
    const containsWebConf = webConferencingRegexCheck.test(lowerCaseBody);

    if (!containsWebConf) {
        asyncResult.asyncContext.completed({ allowEvent: true });
    } else {
        confirmationAndAddCoordinator(asyncResult.asyncContext);
    }
}

function check3Location(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const location = asyncResult.value;
        const containsWebConfInLocation = webConferencingRegexCheck.test(location.toLowerCase());

        if (!containsWebConfInLocation) {
            mailboxItem.body.getAsync('html', { asyncContext: asyncResult.asyncContext }, check4Body);
        } else {
            confirmationAndAddCoordinator(asyncResult.asyncContext);
        }
    } else {
        console.error(asyncResult.error);
    }
}

function check2OptionalAttendees(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const optionalAttendees = asyncResult.value;
        const containsGong = optionalAttendees.some(function (email) {
            return email.emailAddress === GONG_COORDINATOR.EMAIL
        });
        containsExternalAttendee = containsExternalAttendee || optionalAttendees.some(function (email) {
            return !linkedinDomainsRegexCheck.test(email.emailAddress)
        });

        if (!containsGong && containsExternalAttendee) {
            // check web conf in location
            mailboxItem.location.getAsync({ asyncContext: asyncResult.asyncContext }, check3Location);
        } else {
            asyncResult.asyncContext.completed({ allowEvent: true });
        }
    } else {
        console.error(asyncResult.error);
    }
}

function check1RequiredAttendees(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const requiredAttendees = asyncResult.value;
        containsExternalAttendee = requiredAttendees.some(function (email) {
            return !linkedinDomainsRegexCheck.test(email.emailAddress)
        });

        const containsGong = requiredAttendees.some(function (email) {
            return email.emailAddress === GONG_COORDINATOR.EMAIL
        });

        if (!containsGong) {
            mailboxItem.optionalAttendees.getAsync({ asyncContext: asyncResult.asyncContext }, check2OptionalAttendees);
        } else {
            // sendEvent()
            asyncResult.asyncContext.completed({ allowEvent: true });
        }
    } else {
        console.error(asyncResult.error);
    }
}

function itemSendHandler(event) {
    mailboxItem.requiredAttendees.getAsync({ asyncContext: event }, check1RequiredAttendees);
    // add all function here
}
