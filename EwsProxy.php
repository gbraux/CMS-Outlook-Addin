<?php
/*
--- Cisco CMS Outlook Addin /w OBTP ---

EwsProxy.php : Server side PHP script to set the UCCapabilities property of a calendar Item through EWS (Exchange Web Services)

Initial Creator : Guillaume BRAUX (gubraux@cisco.com)
Released under the GNU General Public License v3
*/

include ("php-ews/ExchangeWebServices.php");
include ("php-ews/EWS_Exception.php");
include ("php-ews/EWSType.php");
include ("php-ews/NTLMSoapClient.php");
include ("php-ews/EWSType/FindItemType.php");
include ("php-ews/EWSType/FindFolderType.php");
include ("php-ews/EWSType/FolderQueryTraversalType.php");
include ("php-ews/EWSType/FolderResponseShapeType.php");
include ("php-ews/EWSType/DefaultShapeNamesType.php");
include ("php-ews/EWSType/ItemResponseShapeType.php");
include ("php-ews/EWSType/NonEmptyArrayOfBaseFolderIdsType.php");
include ("php-ews/EWSType/DistinguishedFolderIdType.php");
include ("php-ews/EWSType/DistinguishedFolderIdNameType.php");
include ("php-ews/EWSType/IndexedPageViewType.php");
include ("php-ews/NTLMSoapClient/Exchange.php");
include ("php-ews/EWSType/ContactsViewType.php");
include ("php-ews/EWSType/ItemQueryTraversalType.php");
include ("php-ews/EWSType/RestrictionType.php");
include ("php-ews/EWSType/ExistsType.php");
include ("php-ews/EWSType/PathToExtendedFieldType.php");
include ("php-ews/EWSType/DistinguishedPropertySetIdType.php");
include ("php-ews/EWSType/MapiPropertyTypeType.php");
include ("php-ews/EWSType/ContainsExpressionType.php");
include ("php-ews/EWSType/PathToUnindexedFieldType.php");
include ("php-ews/EWSType/ConstantValueType.php");
include ("php-ews/EWSType/ContainmentComparisonType.php");
include ("php-ews/EWSType/ContainmentModeType.php");
include ("php-ews/EWSType/NonEmptyArrayOfFieldOrdersType.php");
include ("php-ews/EWSType/FieldOrderType.php");
include ("php-ews/EWSType/ExchangeImpersonationType.php");
include ("php-ews/EWSType/ConnectingSIDType.php");
include ("php-ews/EWSType/UpdateItemType.php");
include ("php-ews/EWSType/ExtendedPropertyType.php");
include ("php-ews/EWSType/NonEmptyArrayOfBaseItemIdsType.php");
include ("php-ews/EWSType/ItemIdType.php");
include ("php-ews/EWSType/ItemChangeType.php");
include ("php-ews/EWSType/NonEmptyArrayOfItemChangeDescriptionsType.php");
include ("php-ews/EWSType/ItemType.php");
include ("php-ews/EWSType/SetItemFieldType.php");



// CONFIG -------------------------------------
//error_reporting(0);
include ("config.php");
// -------------------------------------------


// Wait 5 seconds to ensure that the draft calendar item is saved by the Outlook client to the Server
sleep(5);

// Get Data from Outlook Add-In
$addin_id = $_GET['addin_id'];
$prop_guid = $_GET['prop_guid'];
$dest_uri = $_GET['dest_uri'];
$email = $_GET['email'];

//Generate UCCapabilities property to be added to the calendar (mimic Webex PTools for OBTP)
$webex_prop = '<?xml version="1.0"?><CiscoOI><PTVersion>310000</PTVersion><PTReleaseVersion>31.5.1.60</PTReleaseVersion><OIVersion><CreatorOS>Windows</CreatorOS><ClientOS>Windows</ClientOS></OIVersion><PTFeatureConfig>1</PTFeatureConfig><ExternalSIPUrl>'.$dest_uri.'</ExternalSIPUrl><WebExOI>BLABLA</WebExOI><WebExSegmentID>PHNlZz48dHlwZT5ib29rbWFyazwvdHlwZT48cGF0dGVybj48Y2F0ZWdvcnk+cGxhaW5UZXh0PC9jYXRlZ29yeT48bmFtZT5XQlg2RjUxRTwvbmFtZT48dmVyaWZ5Q29kZT44MTE4MDwvdmVyaWZ5Q29kZT48L3BhdHRlcm4+PC9zZWc+DQoAAA==</WebExSegmentID><WebEx><Product><Major>Train</Major><Minor>T29</Minor><SP>8</SP><EP></EP><OtherFlag></OtherFlag></Product><MeetingInfo><site>acecloud.webex.com</site><brandName>acecloud</brandName><LoginName>gubraux</LoginName><LoginAccount>gubraux@cisco.com</LoginAccount><HostName>gubraux</HostName><HostAccount>gubraux@cisco.com</HostAccount><HostID>488092317</HostID><svcType>MC</svcType><MeetingKey>201430513</MeetingKey><audioType>2</audioType><meetingType>3</meetingType><meetingTemplateKey>S;MC;en_US;9.1;1445782;MC Default;D; ;</meetingTemplateKey><EmailBody><Version>1</Version><TagDataLength>100</TagDataLength><BeginTag></BeginTag><EndTag></EndTag></EmailBody></MeetingInfo></WebEx><WebExPMR>AAA=</WebExPMR></CiscoOI>';


// ------------- GET CALENDAR PROPERTIES & ITEM_ID -------------
// Use the custom GUID set by the Addin to find the EWS Item_ID

// UGLY HACK BELLOW : The UCCapabilities property CANNOT be set on the calendar item BEFORE it is sent by the user.
// (because all server-side properties that could have been customized will always overwritten by the send item)
// As we need to write the property AFTER the calendar is sent, we will monitor (every 5 seconds) the item until we get MeetingRequestWasSent == true, 
// and we can them write the property. Monitoring time is limited to the PHP session timer, ie. 5 minutes (at least with default IIS7 + PHP5 config)
// Yes, this is ugly, and you are f***ed-up if you wait more than 5 min before sending your meeting request !
// There may be somthing to to with EWS notifications ...

$ews = new ExchangeWebServices($ews_server, $ews_admin_username, $ews_admin_password);

$ei = new EWSType_ExchangeImpersonationType();
$sid = new EWSType_ConnectingSIDType();
$sid->PrimarySmtpAddress = $email;
$ei->ConnectingSID = $sid;
$ews->setImpersonation($ei);

$isSent = 0;
$calendar_id = "";
$calendar_changekey = "";


while ($isSent != 1)
{
$request = new EWSType_FindItemType();

$request->ItemShape = new EWSType_ItemResponseShapeType();
$request->ItemShape->BaseShape = EWSType_DefaultShapeNamesType::ALL_PROPERTIES;

$request->Traversal = EWSType_ItemQueryTraversalType::SHALLOW;

$request->ParentFolderIds = new EWSType_NonEmptyArrayOfBaseFolderIdsType();
$request->ParentFolderIds->DistinguishedFolderId = new EWSType_DistinguishedFolderIdType();
$request->ParentFolderIds->DistinguishedFolderId->Id = EWSType_DistinguishedFolderIdNameType::CALENDAR;

$request->Restriction = new EWSType_RestrictionType();
$request->Restriction->Contains = new EWSType_ContainsExpressionType();
$request->Restriction->Contains->ExtendedFieldURI = new EWSType_PathToExtendedFieldType();
$request->Restriction->Contains->ExtendedFieldURI->DistinguishedPropertySetId = new EWSType_DistinguishedPropertySetIdType();
$request->Restriction->Contains->ExtendedFieldURI->DistinguishedPropertySetId->_ = EWSType_DistinguishedPropertySetIdType::PUBLIC_STRINGS;
$request->Restriction->Contains->ExtendedFieldURI->PropertyName = 'cecp-'.$addin_id;
$request->Restriction->Contains->ExtendedFieldURI->PropertyType = new EWSType_MapiPropertyTypeType();
$request->Restriction->Contains->ExtendedFieldURI->PropertyType->_ = EWSType_MapiPropertyTypeType::STRING;

$request->Restriction->Contains->Constant = new EWSType_ConstantValueType();
$request->Restriction->Contains->Constant->Value = '{"prop_guid":'.$prop_guid.'}';

$response = $ews->FindItem($request);
echo '<pre>'.print_r($response, true).'</pre>';
file_put_contents("ewslog.txt", print_r($response, true), FILE_APPEND | LOCK_EX);

$isSent = $response->ResponseMessages->FindItemResponseMessage->RootFolder->Items->CalendarItem->MeetingRequestWasSent;
//echo "IS_SENT : ".$isSent;
file_put_contents("ewslog.txt", time()." IS_SENT : ".$isSent, FILE_APPEND | LOCK_EX);

// Got the Item_ID (and Change_ID)
$calendar_id = $response->ResponseMessages->FindItemResponseMessage->RootFolder->Items->CalendarItem->ItemId->Id;
$calendar_changekey = $response->ResponseMessages->FindItemResponseMessage->RootFolder->Items->CalendarItem->ItemId->ChangeKey;
sleep(5);

}





// ----------------- SET THE UCCAPABILITIES PROPERTY TO MIMIC A WEBEX PTOOLS BOOKING

$request = new EWSType_UpdateItemType();

$request->SendMeetingInvitationsOrCancellations = 'SendToAllAndSaveCopy';
$request->MessageDisposition = 'SendAndSaveCopy';
$request->ConflictResolution = 'AlwaysOverwrite';
$request->ItemChanges = array();

// Build out item change request.
$change = new EWSType_ItemChangeType();
$change->ItemId = new EWSType_ItemIdType();
$change->ItemId->Id = $calendar_id;
$change->ItemId->ChangeKey = $calendar_changekey;
$change->Updates = new EWSType_NonEmptyArrayOfItemChangeDescriptionsType();

$change->Updates->SetItemField = array();
$contact = new EWSType_ItemType();

// Build the extended property and set it on the item.
$property = new EWSType_ExtendedPropertyType();
$property->ExtendedFieldURI = new EWSType_PathToExtendedFieldType();
$property->ExtendedFieldURI->PropertyName = 'UCCapabilities';
$property->ExtendedFieldURI->PropertyType = EWSType_MapiPropertyTypeType::STRING;
$property->Value = $webex_prop;
$contact->ExtendedProperty = $property;

// Build the set item field object and set the item on it.
$field = new EWSType_SetItemFieldType();
$field->ExtendedFieldURI = new EWSType_PathToExtendedFieldType();
$field->ExtendedFieldURI->PropertyName = 'UCCapabilities';
$field->ExtendedFieldURI->PropertySetId = '00020329-0000-0000-C000-000000000046';
$field->ExtendedFieldURI->PropertyType = EWSType_MapiPropertyTypeType::STRING;
$field->Contact = $contact;

$change->Updates->SetItemField[] = $field;
$request->ItemChanges[] = $change;

$response = $ews->UpdateItem($request);
var_dump($response);
//file_put_contents("ewslog.txt", print_r($response, true), FILE_APPEND | LOCK_EX);

?>