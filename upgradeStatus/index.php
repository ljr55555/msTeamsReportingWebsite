<head><title>Teams Migration State</title>
<link rel="stylesheet" type="text/css" href="..\formatting.css"></head>
<body>

<?php
require_once '../resources/vendor/autoload.php';

use Office365\PHP\Client\Runtime\Auth\AuthenticationContext;
use Office365\PHP\Client\SharePoint\ClientContext;
use Office365\PHP\Client\SharePoint\ListCreationInformation;
use Office365\PHP\Client\SharePoint\SPList;

require_once '../config.php';                           // Site specific connection information

$listTitle = 'TeamsUsageStats';

$strStatDate  = date('Y-m-d', mktime(0, 0, 0, date("m") , date("d") - 3, date("Y")));

$var = @$_GET['strUID'] ;
$strSearchUser = trim($var);
if ( !($strSearchUser == "") ){

    $authCtx = new AuthenticationContext($Url);
    $authCtx->acquireTokenForUser($UserName,$Password);

    $ctx = new ClientContext($Url,$authCtx);
    $web = $ctx->getWeb();

    echo "<table border=1>\n";
    echo "<tr><th>ID</th><th>Interop Mode</th><th>Last Used Teams</th><th>YTD Private Chats</th><th>YTD Teams Chats</ht><th>YTD Meetings</th></tr>\n";

    $strSearchUserWithDomain = $strSearchUser . "@windstream.com";

    // Get Teams upgrade state
    $list = $web->getLists()->getByTitle("Teams Interoperability");
    $items = $list->getItems()->filter("User eq '$strSearchUserWithDomain'");

    $ctx->load($items);
    $ctx->executeQuery();
    $strTeamsState = NULL;
    foreach( $items->getData() as $item ) {
        $strTeamsState = $item->Interoperability_x0020_Level;
    }

    if($strTeamsState == NULL){
        $strTeamsState = "Islands";
    }

    $list = $web->getLists()->getByTitle($listTitle);
    $items = $list->getItems()->filter("Title eq '$strSearchUser' and Active eq '1'");

    $ctx->load($items);
    $ctx->executeQuery();
    $iFound = 0;
    foreach( $items->getData() as $item ) {
        $strExplodedDate = explode('T', $item->DateLastUsedTeams);
        $strDateLastUsed = $strExplodedDate[0];
        $iFound = 1;
        echo "<tr><td>{$item->Title}</td><td>$strTeamsState</td><td>$strDateLastUsed</td><td>{$item->eb7w}</td><td>{$item->t0gh}</td><td>{$item->s4xd}</td></tr>\n";

    }
    if($iFound != 1){
        echo "<td>$strSearchUser</td><td>$strTeamsState</td><td colspan=4>-- No Teams Usage --</td></tr>\n";
    }


    print "</table><br>\n";
    echo "<P><font>Teams interop modes can be <a href=\"https://windstream.jiveon.com/community/information-technology/it-essentials-for-employees/blog/2019/03/07/teams-and-skype-interoperability/#islandMode\">Islands</a> or <a href=\"https://windstream.jiveon.com/community/information-technology/it-essentials-for-employees/blog/2019/03/07/teams-and-skype-interoperability/#teamsOnlyMode\">UpgradeToTeams</a></font></P>\n";
}

echo "<P><font size=\"+1\">Enter an ID to see the user's Teams interoperability mode and info on their Teams usage. </font></P>\n";

echo "<form name=\"form\" action=\"index.php\" method=\"get\">";
echo "<input type=\"text\" name=\"strUID\" value=\"$strSearchUser\"/>&nbsp;&nbsp;&nbsp;<input type=\"submit\" name=\"Submit\" value=\"Search\" />";
echo "</form>";
echo "<BR>";

    echo "</body>";
echo "</html>";

?>
