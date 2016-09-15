<?php
/*
--- Cisco CMS Outlook Addin /w OBTP ---

CmsProxy.php : Server side PHP script to make REST requests to the CMS Server API, and get the default space details of a user

Initial Creator : Guillaume BRAUX (gubraux@cisco.com)
Released under the GNU General Public License v3
*/



// CONFIG -------------------------------------
include ("config.php");
error_reporting(0);
// -------------------------------------------




// Get (from URL param) a portion of the username to search for it's default Space
$userFilter = $_GET['userFilter'];

// ---------------------- SEARCH FOR CMS USER ID ---------------------

$request = $cms_api_base_url . "users?filter=" . $userFilter;

//Start the Curl session
$session = curl_init($request);

curl_setopt($session, CURLOPT_HEADER, ($headers == "true") ? true : false);
curl_setopt($session, CURLOPT_FOLLOWLOCATION, true);
curl_setopt($session, CURLOPT_SSL_VERIFYPEER, false);
curl_setopt($session, CURLOPT_RETURNTRANSFER, true);
curl_setopt($session, CURLOPT_USERPWD, $cms_admin_username . ":" . $cms_admin_password);

$response = curl_exec($session);
$xml = new SimpleXMLElement($response);
$cms_user_id = $xml->user[0]['id'];

curl_close($session);


// ---------------------- GET CMS DEFAULT SPACE ID ---------------------

$request = $cms_api_base_url . "users/".$cms_user_id."/usercoSpaces";

//Start the Curl session
$session = curl_init($request);

curl_setopt($session, CURLOPT_HEADER, ($headers == "true") ? true : false);
curl_setopt($session, CURLOPT_FOLLOWLOCATION, true);
curl_setopt($session, CURLOPT_SSL_VERIFYPEER, false);
curl_setopt($session, CURLOPT_RETURNTRANSFER, true);
curl_setopt($session, CURLOPT_USERPWD, $cms_admin_username . ":" . $cms_admin_password);

$response = curl_exec($session);

$xml = new SimpleXMLElement($response);
$cms_cospace_id = $xml->userCoSpace[0]['id'];

curl_close($session);


// ---------------------- GET CMS SPACE DETAILS ---------------------

$request = $cms_api_base_url . "cospaces/".$cms_cospace_id;

//Start the Curl session
$session = curl_init($request);

curl_setopt($session, CURLOPT_HEADER, ($headers == "true") ? true : false);
curl_setopt($session, CURLOPT_FOLLOWLOCATION, true);
curl_setopt($session, CURLOPT_SSL_VERIFYPEER, false);
curl_setopt($session, CURLOPT_RETURNTRANSFER, true);
curl_setopt($session, CURLOPT_USERPWD, $cms_admin_username . ":" . $cms_admin_password);

$response = curl_exec($session);

$xml = new SimpleXMLElement($response);

//Extract CMS Space details from API XML Answer
$cms_cospace_name = (string)$xml->name[0];
$cms_cospace_uri = (string)$xml->uri[0];
$cms_cospace_dn = (string)$xml->callId[0];
if ($xml->passcode[0] == null)
    $cms_cospace_pin = null;
else
    $cms_cospace_pin = (string)$xml->passcode[0];
$cms_cospace_secret = (string)$xml->secret[0];

$cms_cospace_webrtc = $cms_webrtc_base_url . $cms_cospace_dn."&secret=".$cms_cospace_secret;

// Build an array containing the CMS Space details
$cms_cospace_array = array(
   "cms_cospace_name" => $cms_cospace_name,
   "cms_cospace_uri" => $cms_cospace_uri.$sip_domain,
   "cms_cospace_dn" => $cms_cospace_dn,
   "cms_cospace_pin" => $cms_cospace_pin,
   "cms_cospace_webrtc" => $cms_cospace_webrtc,
   "cms_phone_sda" => $phone_sda
);

// Write the array in JSON (retreived by Addin JS)
echo json_encode($cms_cospace_array);

curl_close($session);
?>