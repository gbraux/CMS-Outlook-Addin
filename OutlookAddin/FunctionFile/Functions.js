/*
--- Cisco CMS Outlook Addin /w OBTP ---

Funtion.js : Main script run by Outlook (client side) when clicking the addin button

Initial Creator : Guillaume BRAUX (gubraux@cisco.com)
Released under the GNU General Public License v3
*/


// ------------------ CONFIG ----------------------------
// Note : Various labels & text content are still in the code bellow and need to be changed (language ...)

//Outlook addin ID as defined in the addin XML Manifest
var addin_id = "1e5a160d-61bd-49c9-9936-49999999999d"
//-------------------------------------------------------

var item;
var cospace_uri;
var body_type;

Office.initialize = function () {
  item = Office.context.mailbox.item;

};

function showMessage(message, icon, event) {
  Office.context.mailbox.item.notificationMessages.replaceAsync('msg', {
    type: 'informationalMessage',
    icon: icon,
    message: message,
    persistent: false
  }, function (result) {
    event.completed();
  });
}

function setMeeting(event) {
  meetingSchedRoutine(event);
}

// Main Routine launched when the "Add CMS Meeting" is pressed
function meetingSchedRoutine(event) {
  item.body.getTypeAsync(
    function (result) {
      if (result.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
      }
        
        // We keep the body formating (HTML vs Plain Text)
        body_type = result.value;


          // Generate user descriptor that can be used in CMS user search filter.
          // It is a hack if the CMS username is not (exactly) the email address of the outlook user (ie. Outlook : gubraux@ciscofrance.com | CMS : gubraux.cms@ciscofrance.com)
          // Here, we only keep the left portion of the email

          userFilter = Office.context.mailbox.userProfile.emailAddress;
          userFilter = userFilter.substring(0, userFilter.indexOf("@"));
          // showMessage("User = " + userFilter, 'icon-16', event);

          // DEBUG - Force Username ------

          //userFilter = "pierre";

          // -----------------------------


          // Request personnal room details from CMS API through server side proxy call
          // Proxy is mandatory as Javascript is not able to do Cross Domain requests (Outlook addin CANNOT make REST/AJAX requests to other servers than defined in the manifest)
          // So the addin can't reach CMS directly
          var settings = {
            "async": false,
            "crossDomain": false,
            "url": "https://showroom.ciscofrance.com/bookingplugin/CmsProxy.php?userFilter=" + userFilter,
            "method": "GET"
          }

          $.ajax(settings).done(function (response) {
            json = JSON.parse(response);

            // Check if PIN code is defined in CMS
            pin = "None"
            if (json.cms_cospace_pin != null)
              pin = json.cms_cospace_pin;

            cospace_uri = json.cms_cospace_uri;

            // Generate text to be appened in the body of the Outlook invite
            // Formating can be HTML or plain text. We only manage HTML.
            if (result.value == Office.MailboxEnums.BodyType.Html)
            {
              inviteText = "<div style='font-size:10.0pt;font-family:\"Segoe UI\",sans-serif;mso-fareast-font-family:\"Times New Roman\"'><br> --- " + Office.context.mailbox.userProfile.displayName + " invites you to this virtual meeting (" + json.cms_cospace_name + ") --- <br><br> You have several ways to join this meeting : <br><br> <ul> <li>From a <b>Computer (PC/Mac)</b> or <b>a Smartphone/Tablet (iOS-Android)</b>, click the following link : <a href=\"" + json.cms_cospace_webrtc + "\">" + json.cms_cospace_webrtc + "</a></li> <li>From a standard-based <b>videoconferencing endpoint</b> (SIP/H.323), enter the following video address (with your remote or touch panel) : " + json.cms_cospace_uri + "</li> <li>From a <b>Unified Communication client</b> (ie. Cisco Jabber, Microsoft Skype for Business), enter or click the following URI : <a href=\"sip:" + json.cms_cospace_uri + "\">sip:" + json.cms_cospace_uri + "</a></li> <li>From a <b>phone</b>, dial " + json.cms_phone_sda + ", and enter the meeting ID (" + json.cms_cospace_dn + ") </li></ul>Meeting PIN : " + pin + "<br><br><b>Note :</b> If you are near a <b>Proximity enabled Cisco video endpoint</b>, you can <a href=\"proximity:" + json.cms_cospace_uri + "\">click here</a> to connect the endpoint to the meeting using your Smartphone<br></div>";
            }
            else
            {
              // Plain text
            }

            if (json.cms_cospace_pin != null)
              pin = json.cms_cospace_pin;


            // Append meeting details to body  
            item.body.getAsync(Office.CoercionType.Html,{ asyncContext:"This is passed to the callback" },function (asyncResult) {
                if (asyncResult.status ==
                  Office.AsyncResultStatus.Failed) {
                  write(asyncResult.error.message);
                }
                else {

                  bdy = asyncResult.value;
                  bdy = bdy.replace("</body>",inviteText+"</body>");

                  item.body.setAsync(bdy, { coercionType:"html", asyncContext:"This is passed to the callback" }, function callback(result) {


                  // Write additionnal details in the meeting request Location field
                  item.location.setAsync("CMS Virtual Meeting (ID : " + json.cms_cospace_dn + ")", function (result) {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                      Office.context.mailbox.item.notificationMessages.addAsync('setSubjectError', {
                        type: 'errorMessage',
                        message: 'Failed to set subject: ' + result.error
                      });

                      event.completed();
                    }
                  });


                  // Write custom property (generated GUID) to the invite (the GUID will allow our EWS script to find the calendar item server-side by searching for this GUID)
                  item.loadCustomPropertiesAsync(function (result) {
                    var guid = generateGuid();
                    _customProps = result.value;
                    _customProps.set("prop_guid", guid);

                    _customProps.saveAsync(function (result) {

                      // Save the meeting request draft to Exchange (same as clicking on the "save" button). Will allow our EWS server-side script to find and edit it.
                      item.saveAsync(function (result) {


                        // Call to Exchange EWS (through Proxy) to create and populate the "UCCapabilites" property into the server stored calendar item.
                        // UCCapabilites is (normaly) used by Webex Ptoos & TMS-XE to get OBTP when scheduling a CMR Cloud meeting
                        // It does not wait for an answer.
                        var settings = {
                          "async": true,
                          "crossDomain": false,
                          "url": "https://showroom.ciscofrance.com/bookingplugin/EwsProxy.php?addin_id=" + addin_id + "&prop_guid=" + guid + "&email=" + Office.context.mailbox.userProfile.emailAddress + "&dest_uri=" + cospace_uri,
                          "method": "GET"
                        }
                          $.ajax(settings);
                          showMessage("Meeting details have been addded successfuly", 'icon-16', event);

                            

                          });
                      });
                    });
                  });
                }
              });
          });
    });
}

//Writes to a div with id='message' on the page.
function write(message) {
  document.getElementById('message').innerText += message;
}

function generateGuid() {
  var randomnumber = Math.floor(Math.random() * 10000000000000001)
  return randomnumber
}

function sleep(milliseconds) {
  var start = new Date().getTime();
  for (var i = 0; i < 1e7; i++) {
    if ((new Date().getTime() - start) > milliseconds) {
      break;
    }
  }
}