/* global Office */
const linkedinDomains = new Array('linkedin\\.com$', 'linkedin\\.biz$', 'glintinc\\.com$');
const linkedinDomainsRegex = new RegExp(linkedinDomains.join('|'));
const GONG_COORDINATOR = { NAME: 'Gong Coordinator', EMAIL: 'coordinator@inbound.gong.io' };
const ITEM_TYPE = {
    APPOINTMENT: 'appointment',
}
const LOCAL_STORAGE = {
    KEY: 'linkedin-add-in_always_invite_gong',
    TRUE_VALUE: 'true',
}
const URL = {
    ADD_IN_DOMAIN: 'https://gong-for-linkedin.s3.amazonaws.com/',
    SCHEDULING_DIALOG: 'scheduling-add-in-dialog.html'
}
// A Flag to indicate if there is External Attendee
let containsExternalAttendee;
let mailboxItem;

Office.initialize = function () {
    mailboxItem = Office.context.mailbox.item;
};

function itemSendHandler(event) {
    if (mailboxItem.itemType === ITEM_TYPE.APPOINTMENT) {
        mailboxItem.requiredAttendees.getAsync({ asyncContext: event }, checkRequiredAttendees);

        function checkRequiredAttendees(requiredAttendeesResult) {
            if (requiredAttendeesResult.status === Office.AsyncResultStatus.Succeeded) {

                containsExternalAttendee = isExternalAttendeeInvited(requiredAttendeesResult.value);
                if (!isGongCoordinatorInvited(requiredAttendeesResult.value)) {
                    mailboxItem.optionalAttendees.getAsync({ asyncContext: requiredAttendeesResult.asyncContext }, checkOptionalAttendees);
                } else {
                    sendInvite(requiredAttendeesResult.asyncContext);
                }
            } else {
                console.error(requiredAttendeesResult.error);
            }
        }

        function checkOptionalAttendees(optionalAttendeesResult) {
            if (optionalAttendeesResult.status === Office.AsyncResultStatus.Succeeded) {

                if (!isGongCoordinatorInvited(optionalAttendeesResult.value) && (containsExternalAttendee || isExternalAttendeeInvited(optionalAttendeesResult.value))) {
                    confirmationAndAddCoordinator(optionalAttendeesResult.asyncContext);
                } else {
                    sendInvite(optionalAttendeesResult.asyncContext);
                }
            } else {
                console.error(optionalAttendeesResult.error);
            }
        }

        function confirmationAndAddCoordinator(event) {
            if (alwaysAddGongToInvite()) {
                addGongCoordinator(event);
                sendInvite(event);
            } else {
                openDialog(event);
            }
        }

        function alwaysAddGongToInvite() {
            const addGongAutomatically = localStorage.getItem(LOCAL_STORAGE.KEY);
            return addGongAutomatically != null && addGongAutomatically === LOCAL_STORAGE.TRUE_VALUE;
        }

        function openDialog(event) {
            Office.context.ui.displayDialogAsync(URL.ADD_IN_DOMAIN + URL.SCHEDULING_DIALOG, {
                height: 60,
                width: 30,
                displayInIframe: true,
            }, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error(asyncResult.error.message);
                } else {
                    const dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (messageEvent) {
                        dialog.close();
                        // handle message from dialog (scheduling-add-in-dialog.html)
                        if (messageEvent.message) {
                            addGongCoordinator(event);
                        }
                        sendInvite(event);
                    });
                    dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
                        dialog.close();
                        sendInvite(event);
                    });
                }
            });
        }

        function addGongCoordinator(event) {
            mailboxItem.optionalAttendees.addAsync([{
                displayName: GONG_COORDINATOR.NAME,
                emailAddress: GONG_COORDINATOR.EMAIL,
            }], { asyncContext: event });
        }

        function sendInvite(event) {
            event.completed({ allowEvent: true });
        }

        function isExternalAttendeeInvited(attendees) {
            return attendees.some(function (attendee) {
                return !linkedinDomainsRegex.test(attendee.emailAddress)
            });
        }

        function isGongCoordinatorInvited(attendees) {
            return attendees.some(function (attendee) {
                return attendee.emailAddress === GONG_COORDINATOR.EMAIL
            });
        }
    } else {
        event.completed({ allowEvent: true });
    }

}
