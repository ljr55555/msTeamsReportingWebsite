<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="utf-8" />
    <title>Your Organization's Teams Adoption Status</title>
    <link rel="stylesheet" type="text/css" href="..\formatting.css">
    <link rel="stylesheet" type="text/css" href="..\Resources\datatables\dataTables.css">

    <script src="/resources/jquery.js"></script>
    <script type="text/javascript" charset="utf8" src="..\Resources\datatables\dataTables.js"></script>
    <script type="text/javascript">
    $(document).ready( function () {
        $('#teamsUsageTable').DataTable({
            paging: false
        });
    } );
    </script>

</head>
<body>

<?php
//error_reporting(E_ALL);       // for debugging
set_time_limit(30000);          // required or high ranking people never see their data

if( ! isset($_POST['MaxRecusion'])){
	$iMaxRecursionLevels = 1;
}
else{
	$iMaxRecursionLevels = $_POST['MaxRecusion'];
}


// resources
require_once '..\resources\vendor\autoload.php';

use Office365\PHP\Client\Runtime\Auth\AuthenticationContext;
use Office365\PHP\Client\SharePoint\ClientContext;
use Office365\PHP\Client\SharePoint\ListCreationInformation;
use Office365\PHP\Client\SharePoint\SPList;

require_once '../config.php';                           // Site specific connection information


// Recursive function to find subtree org structure
function getDirectReports($ctxClientConnection, $webSPO, $ldapDS, $ldapRoot, $strRecurseFQDN, &$strTableDataArray, $iRecursionLevel, $iMaxLevels){
	//echo "<P>I am on recursion level $iRecursionLevel of max $iMaxLevels</P>\n";
    if($iRecursionLevel > $iMaxLevels){
		return;
	}
	else{
		$iRecursionLevel = $iRecursionLevel + 1;
		// find all direct reports
		$scriteria2="(&(manager=" . str_replace("\\","\\\\",$strRecurseFQDN) . ")(|(sAMAccountName=e0*)(sAMAccountName=n9*)))";

		$sr2=ldap_search($ldapDS,$ldapRoot,$scriteria2);
		$infoReports = ldap_get_entries($ldapDS, $sr2);

		// for each direct report, get Teams usage stats
		for ($i=0; $i<$infoReports["count"]; $i++) {
			$strUserRecord = getUsageStats($ctxClientConnection, $webSPO, $infoReports[$i]["samaccountname"][0], $infoReports[$i]["displayname"][0]);
			array_push($strTableDataArray,$strUserRecord);
			getDirectReports($ctxClientConnection, $webSPO, $ldapDS, $ldapRoot, $infoReports[$i]["distinguishedname"][0], $strTableDataArray, $iRecursionLevel, $iMaxLevels);
		}
	}
}
// Get Teams usage for individual user ($strReportUID)
function getUsageStats($ctxClientConnection, $webSPO, $strReportUID, $strReportDisplayName){
    $strSearchUserWithDomain = $strReportUID . "@windstream.com";
    // Get Teams upgrade state
    $list = $webSPO->getLists()->getByTitle("Teams Interoperability");
    $items = $list->getItems()->filter("User eq '$strSearchUserWithDomain'");

    $ctxClientConnection->load($items);
    $ctxClientConnection->executeQuery();
    $strTeamsState = NULL;
    foreach( $items->getData() as $item ) {
        $strTeamsState = $item->Interoperability_x0020_Level;
    }

    if($strTeamsState == NULL){
        $strTeamsState = "Islands";
    }
    
    if( strcmp("UpgradeToTeams",$strTeamsState) == 0){
        $strTeamsState = "Upgraded To Teams";
    }


    $strRecord = "<tr><td>$strReportDisplayName</td><td>{$strReportUID}</td><td>{$strTeamsState}</td>";

    $list = $webSPO->getLists()->getByTitle("TeamsUsageStats");
    $items = $list->getItems()->filter("Title eq '$strReportUID' and Active eq '1'");

    $ctxClientConnection->load($items);
    $ctxClientConnection->executeQuery();
    $iFound = 0;
    foreach( $items->getData() as $item ) {
        $strExplodedDate = explode('T', $item->DateLastUsedTeams);
        $strDateLastUsed = $strExplodedDate[0];
        $iFound = 1;
        // call counts are inaccurate, so removing from report
        // $strRecord = $strRecord .  "<td>$strDateLastUsed</td><td>{$item->eb7w}</td><td>{$item->t0gh}</td><td>{$item->s4xd}</td><td>{$item->rjbz}</td>";
        $strRecord = $strRecord .  "<td>$strDateLastUsed</td><td>{$item->eb7w}</td><td>{$item->t0gh}</td><td>{$item->s4xd}</td>";

    }
    if($iFound == 1){
        $strRecord = $strRecord . "</tr>\n";
    }
    else{
        $strRecord = $strRecord . "<td>-- No Teams Usage --</td><td>-</td><td>-</td><td>-</td></tr>\n";
    }
    return $strRecord;
}

//echo "<P>Started: " .  date(DATE_RFC2822) . "</P>\n";

$strAuthUser = $_SERVER['AUTH_USER'];

if( strpos($strAuthUser, '@windstream.com') !== false ){
    list($strDomain, $strLogonUser) = explode('@', $strAuthUser);

    if($strLogonUser == ""){
        $strLogonUser = $strAuthUser;
    }
    $strLogonUser = $strLogonUser;
}
elseif( strpos($strAuthUser, '\\') !== false ){
    list($strDomain, $strLogonUser) = explode('\\', $strAuthUser);

    if($strLogonUser == ""){
        $strLogonUser = $strAuthUser;
    }
    $strLogonUser = $strLogonUser;
}
else{
    $strLogonUser = $strAuthUser;
}

$authCtx = new AuthenticationContext($Url);
$authCtx->acquireTokenForUser($UserName,$Password);

$ctx = new ClientContext($Url,$authCtx);
$web = $ctx->getWeb();

$ds = ldap_connect($strADDC) or die("Could not connect to LDAP server.");
if ($ds) {
    $ldapbind = ldap_bind($ds, $ldaprdn, $ldappass) or die("Could not bind to LDAP server");
    $strTableData = array();

    $scriteria="(&(sAMAccountName=$strLogonUser))";
    $sr=ldap_search($ds,$ldaproot,$scriteria) or die("$_");

    $infoManager = ldap_get_entries($ds, $sr);
    $strManagerFQDN = $infoManager[0]["distinguishedname"][0];
    $strManagerFirstName = $infoManager[0]["givenname"][0];

	//echo "<P>For $strLogonUser, max recursion level is $iMaxRecursionLevels</P>\n";
    $strUserRecord = getUsageStats($ctx, $web, $infoManager[0]["samaccountname"][0], $infoManager[0]["displayname"][0]);
    array_push($strTableData,$strUserRecord);

    getDirectReports($ctx, $web, $ds, $ldaproot, $strManagerFQDN, $strTableData,1, $iMaxRecursionLevels);
}

$strStatDate  = date('Y-m-d', mktime(0, 0, 0, date("m") , date("d") - 3, date("Y")));   // stats run three days behind, so calulating "stats as of" date for page text

sort($strTableData);
if($iMaxRecursionLevels < 2){
	echo "<P><font size=\"+1\">Happy " . date("l") . ", $strManagerFirstName! Below you will find Teams usage statistics, as of $strStatDate, for your direct reports (if you have any reports!).</font></P>\n";
}
else{
	echo "<P><font size=\"+1\">Happy " . date("l") . ", $strManagerFirstName! Below you will find Teams usage statistics, as of $strStatDate, for $iMaxRecursionLevels levels of the organization that reports up through you (if you have any reports!).</font></P>\n";
}
echo "<table border=0 id=\"teamsUsageTable\" class=\"display\">\n";
echo "<thead><tr><th>Name</th><th>ID</th><th>Interop State</th><th>Last Used Teams</th><th>YTD Private Chats</th><th>YTD Teams Chats</th><th>YTD Meetings</th></tr></thead>\n";
echo "<tbody>";
foreach($strTableData as $strTR){
    echo "$strTR\n";
}
echo "</tbody></table>\n";

echo "<BR><P><font size=\"+1\">If you would like to see a larger view of the organization that reports up through you, select the number of sub-reporting levels you wish to display.</font></P><P><i>Warning</i>: For someone with a large sub-organization, higher numbers may take a <b><i>very</i></b> long time to load.</P>\n";
echo "<form method=\"post\" action=\"{$_SERVER['PHP_SELF']}\">\n";
echo "<select name=\"MaxRecusion\">\n";
for ($i = 1; $i <= 5; $i++){

	echo "		<option value=\"$i\"";
	if( $i == $iMaxRecursionLevels){
		echo " selected=\"true\"";
	}
	echo ">$i</option>\n";
}
echo "</select>\n";
echo "   <input type=\"submit\" name=\"submit\" value=\"Update\"><br>\n";
echo "</form>\n";

//echo "<P>Ended: " . date(DATE_RFC2822) . "</P>\n";

?>
</html>
