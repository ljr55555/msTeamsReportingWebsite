<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="utf-8" />
    <title>Windstream Public Team Spaces</title>
    <link rel="stylesheet" type="text/css" href=".\formatting.css">
    <link rel="stylesheet" type="text/css" href=".\Resources\datatables\dataTables.css">
  
    <script src="/resources/jquery.js"></script>
    <script type="text/javascript" charset="utf8" src=".\Resources\datatables\dataTables.js"></script>
    <script type="text/javascript">
    $(document).ready( function () {
        $('#publicTeams').DataTable({
            paging: false

        });
    } );
    </script>

</head>
<body>

<?php
require_once '.\Resources\vendor\autoload.php';
require_once '../config.php';                           // Site specific connection information

use Office365\PHP\Client\Runtime\Auth\AuthenticationContext;
use Office365\PHP\Client\SharePoint\ClientContext;
use Office365\PHP\Client\SharePoint\ListCreationInformation;
use Office365\PHP\Client\SharePoint\SPList;


$strRecord = NULL;

$authCtx = new AuthenticationContext($Url);
$authCtx->acquireTokenForUser($UserName,$Password);

$ctx = new ClientContext($Url,$authCtx);
$web = $ctx->getWeb();

$list = $web->getLists()->getByTitle("TeamsGroups");
$items = $list->getItems()->filter("Visibility eq 'Public'");

$ctx->load($items);
$ctx->executeQuery();
$strTableData = array();

foreach( $items->getData() as $item ) {
    $strRecord = $strRecord .  "<tr><td><a href=\"https://teams.microsoft.com/l/team/conversations/General?groupId={$item->GroupID}&tenantId=2567b4c1-b0ed-40f5-aee3-58d7c5f3e2b2\">{$item->Title}</a></td><td>{$item->TeamDescription}</td></tr>\n";
    array_push($strTableData,$strRecord);
}

echo "<P>I have " . count($strTableData) . " unsorted list items.</P>\n";

sort($strTableData);
echo "<P>I have " . count($strTableData) . " sorted list items.</P>\n";

echo "<table border=0 id=\"publicTeams\" class=\"display\">\n";
echo "<thead><tr><th>Team Name</th><th>Description</th></tr></thead>\n";
echo "<tbody>";
foreach($strTableData as $strTR){
    echo "$strTR\n";
}
print "</tbody></table>\n";
?>
</html>
