<!DOCTYPE html>
<meta charset="utf-8">
<head>
<link rel="stylesheet" href="usageReporting.css">
<title>Microsoft Teams Usage And Adoption Reports</title>
</head>
<body>

<script src="/resources/Chart.js/node_modules/chart.js/dist/Chart.js"></script>

 <!-- Tab links -->
<div class="tab">
  <button class="tablinks" onclick="drawChart(event, 'createdByMonth')" id="defaultOpen">Group Creation</button>
  <button class="tablinks" onclick="drawChart(event, 'groupVisibility')">Visibility</button>
  <button class="tablinks" onclick="drawChart(event, 'memberCount')">Members Count</button>
  <button class="tablinks" onclick="drawChart(event, 'messageAging')">Activity Aging</button>
  <button class="tablinks" onclick="drawChart(event, 'messageCounts')">YTD Activity By Type</button>
  <button class="tablinks" onclick="drawChart(event, 'SkypeVSTeams')">Skype v/s Teams</button>
  <button class="tablinks" onclick="drawChart(event, 'privateChatLine')">Private Chats / Day</button>
  <button class="tablinks" onclick="drawChart(event, 'teamChatLine')">Team Conversations / Day</button>
  <button class="tablinks" onclick="drawChart(event, 'callsLine')">Calls / Day</button>
  <button class="tablinks" onclick="drawChart(event, 'meetingsLine')">Meetings / Day</button>
  <button class="tablinks" onclick="drawChart(event, 'upgradesLine')">Teams Upgrades</button>
</div>

<!-- Tab content -->
<div id="createdByMonth" class="tabcontent">
	<canvas id="chartGroupsCreatedByMonth" width="805" height="325" />
</div>

<div id="groupVisibility" class="tabcontent">
	<canvas id="chartVisibility" width="805" height="325" />
</div>

<div id="memberCount" class="tabcontent">
        <canvas id="chartMemberCount" width="805" height="325" />
</div>

<div id="messageAging" class="tabcontent">
        <canvas id="chartMessageAging" width="805" height="325" />
</div>

<div id="messageCounts" class="tabcontent" />
	<canvas id="chartMessageCounts" width="805" height="325" />
</div>

<div id="privateChatLine" class="tabcontent" />
	<canvas id="chartPrivateChats" width="805" height="325" />
</div>

<div id="teamChatLine" class="tabcontent" /> 
	<canvas id="chartTeamChats" width="805" height="325" />
</div>

<div id="callsLine" class="tabcontent" />
	<canvas id="chartCalls" width="805" height="325" />
</div>

<div id="meetingsLine" class="tabcontent" />
	<canvas id="chartMeetings" width="805" height="325" />
</div>

<div id="upgradesLine" class="tabcontent" />
	<canvas id="chartUpgrades" width="805" height="325" />
</div>

<div id="SkypeVSTeams" class="tabcontent" />
        <canvas id="chartSkypeVSTeams" width="805" height="325" />
</div>

<script>
document.getElementById("defaultOpen").click();
Chart.defaults.global.defaultColor = 'rgba(255,255,255,1)';

<?php
	error_reporting(0);       // for debugging

	function printRainbowRGBA($iAlpha){
	        echo "                                  'rgba(214, 10, 17, $iAlpha)',\n";
	        echo "                                  'rgba(225, 105, 0, $iAlpha)',\n";
	        echo "                                  'rgba(227, 234, 7, $iAlpha)',\n";
	        echo "                                  'rgba(16, 170, 8, $iAlpha)',\n";
	        echo "                                  'rgba(25, 25, 255, $iAlpha)',\n";
	        echo "                                  'rgba(102, 127, 255, $iAlpha)',\n";
	        echo "                                  'rgba(132, 8, 170, $iAlpha)'\n";
	}
	function printLineChartOptions(){
		echo "                  defaultFontColor: 'white',\n";
		echo "                  legend: {\n";
		echo "                          display: false,\n";
		echo "                          labels: {\n";
		echo "                                  fontColor: 'white'\n";
		echo "                          }\n";
		echo "                  },\n";
		echo "                  elements: {\n";
		echo "                          point: { hitRadius: 5, hoverRadius: 5, radius: 3, pointStyle: 'circle', backgroundColor: 'rgba(255,255,87,0.8)', } \n";
		echo "                  },\n";
		echo "                  tooltips: {\n";
		echo "				callbacks: {\n";
		echo "					label: function(tooltipItem, data) {\n";
		echo "						var tooltipValue = data.datasets[tooltipItem.datasetIndex].data[tooltipItem.index];\n";
		echo "						return parseInt(tooltipValue).toLocaleString();\n";
		echo "					}\n";
		echo "				},\n";
		echo "                          mode: 'point'\n";
		echo "                  },\n";
		echo "                  multiTooltipTemplate : \"<%=datasetLabel%> : <%=value%>\",\n";
		echo "                  events: ['mousemove', 'mouseout', 'click', 'touchstart', 'touchmove'],\n";
		echo "                  scales: {\n";
		echo "                          yAxes: [{\n";
		echo "                                  gridLines: {\n";
		echo "                                          display: true,\n";
		echo "                                          drawBorder: true,\n";
		echo "                                          drawOnChartArea: true,\n";
		echo "                                          zeroLineColor: 'white',\n";
		echo "                                          color: 'rgba(255,255,255,0.7)',\n";
		echo "                                          fontColor: 'white'\n";
		echo "                                  },\n";
		echo "                                  ticks: {\n";
                echo "						callback: function(value, index, values) {\n";
                echo "							return value.toLocaleString();\n";
                echo "						},\n";
		echo "                                          beginAtZero: true,\n";
		echo "                                          fontColor: 'white',\n";
		echo "                                  }\n";
		echo "                          }],\n";
		echo "                          xAxes: [{\n";
		echo "                                  gridLines: {\n";
		echo "                                          display: true,\n";
		echo "                                          drawBorder: true,\n";
		echo "                                          drawOnChartArea: false,\n";
		echo "                                          zeroLineColor: 'white',\n";
		echo "                                          color: 'rgba(255,255,255,0.7)',\n";
		echo "                                          fontColor: 'white'\n";
		echo "                                  },\n";
		echo "                                  ticks: {\n";
		echo "                                          fontColor: 'white',\n";
		echo "                                          autoSkip: true,\n";
		echo "                                          maxTicksLimit: 20\n";
		echo "                                  }\n";
		echo "                          }]\n";
		echo "                  }\n";
	}

        function printPieChartOptions(){
                echo "                  defaultFontColor: 'white',\n";
                echo "                  legend: {\n";
                echo "                          display: true,\n";
		echo "							position: \"top\",\n";
                echo "                          labels: {\n";
                echo "                                  fontColor: 'white'\n";
                echo "                          }\n";
                echo "                  },\n";
                echo "                  elements: {\n";
                echo "                          point: { hitRadius: 5, hoverRadius: 5, radius: 0 } \n";
                echo "                  },\n";
                echo "                  tooltips: {\n";
                echo "                          callbacks: {\n";
		echo "					label: function(tooltipItem, data) {\n";
		echo "						return data.labels[tooltipItem.index] + \": \" + data.datasets[0].data[tooltipItem.index].toLocaleString();\n";
                echo "                                  }\n";
                echo "                          },\n";
                echo "                          mode: 'point'\n";
                echo "                  },\n";
                echo "                  multiTooltipTemplate : \"<%=datasetLabel%> : <%=value%>\",\n";
                echo "                  events: ['mousemove', 'mouseout', 'click', 'touchstart', 'touchmove']\n";
        }

	// resources
	require_once '../resources/vendor/autoload.php';
	require_once '../config.php';				// Site specific connection information

	use Office365\PHP\Client\Runtime\Auth\AuthenticationContext;
	use Office365\PHP\Client\SharePoint\ClientContext;
	use Office365\PHP\Client\SharePoint\ListCreationInformation;
	use Office365\PHP\Client\SharePoint\SPList;
	
	$authCtx = new AuthenticationContext($Url);
	$authCtx->acquireTokenForUser($UserName,$Password);

	$ctxClientConnection = new ClientContext($Url,$authCtx);
	$web = $ctxClientConnection->getWeb();

	$iRecords = 0;

	$list = $web->getLists()->getByTitle("TeamsGroups");
	$iLastRecordCount = $iRecords;
	$items = $list->getItems()->top(10000);

	$ctxClientConnection->load($items);
	$ctxClientConnection->executeQuery();

	$arrayCreatedOn = array();
	$arrayVisibility = array();
	$arrayMembers = array();
	$arrayMsgAging = array();

	foreach( $items->getData() as $item ) {
		$arrayDate = explode("T",$item->CreatedOn);
		$arrayYYYYMM = explode("-",$arrayDate[0]);
		$strCreationYYYYMM = $arrayYYYYMM[0] . $arrayYYYYMM[1];
		$arrayCreatedOn[$strCreationYYYYMM] = $arrayCreatedOn[$strCreationYYYYMM] + 1;
		$arrayVisibility[$item->Visibility] = $arrayVisibility[$item->Visibility] + 1;
		$iMembers = $item->memberCount + $item->externalMemberCount;
		$iMessageAge = $item->lastActivity;

		// Member count: 0, 1, 2-5, 6-10, 11-25, 26-50, 51-100, >100
		if($iMembers == 0)	{	$arrayMembers["0"] = $arrayMembers["0"] + 1;			}
		elseif($iMembers == 1)	{	$arrayMembers["1"] = $arrayMembers["1"] + 1;			}
		elseif($iMembers > 100)	{	$arrayMembers["100Plus"] = $arrayMembers["100Plus"] + 1;	}
		elseif($iMembers >= 51)	{	$arrayMembers["51"] = $arrayMembers["51"] + 1;			}
		elseif($iMembers >= 26)	{	$arrayMembers["26"] = $arrayMembers["26"] + 1;			}
		elseif($iMembers >= 11)	{	$arrayMembers["11"] = $arrayMembers["11"] + 1;			}
		elseif($iMembers >= 6)	{	$arrayMembers["6"] = $arrayMembers["6"] + 1;			}
		elseif($iMembers >= 2)	{	$arrayMembers["2"] = $arrayMembers["2"] + 1;			}

		// Message age: 0, 1-7, 8-14, 15-30, 31-60, 61-180, 181-365, >365
		if($iMessageAge == 0)		{	$arrayMsgAging["0"] = $arrayMsgAging["0"] + 1;		}
		elseif($iMessageAge > 365)	{	$arrayMsgAging["366"] = $arrayMsgAging["366"] + 1;	}
		elseif($iMessageAge >= 181)	{	$arrayMsgAging["181"] = $arrayMsgAging["181"] + 1;	}
		elseif($iMessageAge >= 61)	{	$arrayMsgAging["61"] = $arrayMsgAging["61"] + 1;	}
		elseif($iMessageAge >= 31)	{	$arrayMsgAging["31"] = $arrayMsgAging["31"] + 1;	}
		elseif($iMessageAge >= 15)	{	$arrayMsgAging["15"] = $arrayMsgAging["15"] + 1;	}
		elseif($iMessageAge >= 8)	{	$arrayMsgAging["8"] = $arrayMsgAging["8"] + 1;		}
		elseif($iMessageAge >= 1)	{	$arrayMsgAging["1"] = $arrayMsgAging["1"] + 1;		}
	}


	$strCreationTimeLabels = NULL;
	$strCreationTimeValues = NULL;
	foreach($arrayCreatedOn as $strBucket=>$iValue){
		$strBucket = '"' . mb_substr($strBucket, 0, 4) . "-" . mb_substr($strBucket, 4, 2) . '"';
		if($strCreationTimeLabels){
			$strCreationTimeLabels = $strCreationTimeLabels . ", " . $strBucket;
                        $strCreationTimeValues = $strCreationTimeValues . ", " . $iValue;
		}
		else{
			$strCreationTimeLabels = $strBucket;
			$strCreationTimeValues = $iValue;
		}
	}

	// Groups created by month bar chart
	echo "var ctx = document.getElementById('chartGroupsCreatedByMonth')\n";
	echo "var myChart = new Chart(ctx, \n";
	echo "  {\n";
	echo "		type: 'bar',\n";
	echo "		responsive: true,\n";
	echo "		data: {\n";

	echo "			labels: [$strCreationTimeLabels],\n";
	echo "			datasets: [\n";
	echo "			{\n";
	echo "				fill: true,\n";
	echo "				data: [$strCreationTimeValues],\n";
        echo "				    backgroundColor: 'rgba(84, 118, 242, 0.8)',\n";
        echo "				    borderColor: 'rgba(84, 118, 242, 1)',\n";
        echo "				    borderWidth: 1\n";
	echo "			}],\n";
	echo "	 	 },\n";
	echo "          options: {\n";
	printLineChartOptions();
	echo "          }\n";
	echo "});\n";

	// Group Visibility Pie Chart
	echo "var ctxVisibility = document.getElementById('chartVisibility')\n";
	echo "var myVisibilityPieChart = new Chart(ctxVisibility, {\n";
	echo "	type: 'pie',\n";
	echo "  responsive: true,\n";
	echo "          data: {\n";

	echo "                  labels: ['Public', 'Private'],\n";
	echo "                  datasets: [\n";
	echo "                  {\n";
	echo "                          fill: true,\n";
	echo "                          data: [" . $arrayVisibility['Public'] . ", " . $arrayVisibility['Private'] . "],\n";
	echo "                         	backgroundColor: [\n";
	echo "					'rgba(120, 31, 150, 0.8)',\n";
	echo "					'rgba(201, 127, 225, 0.8)',\n";
	echo "				],\n";
	echo "                         	borderColor: [\n";
	echo "					'rgba(120, 31, 150, 1)',\n";
	echo "					'rgba(201, 127, 225, 1)',\n";
	echo "				],\n";
	echo "                         	borderWidth: 1\n";
	echo "                  }],\n";
	echo "           },\n";
	echo "		options: {\n";
	printPieChartOptions();
	echo "          },\n";
	echo "});\n";


	// Group Member Count Pie Chart
	// Member count: 0, 1, 2-5, 6-10, 11-25, 26-50, 51-100, >100
	$strMemberCountLabels = "'1','2-5','6-10','11-25','26-50','51-100','100Plus'";
	$strMemberCountValues = "{$arrayMembers['1']}, {$arrayMembers['2']}, {$arrayMembers['6']}, {$arrayMembers['11']}, {$arrayMembers['26']}, {$arrayMembers['51']}, {$arrayMembers['100Plus']}, ";
	
	echo "var ctxMemberCount = document.getElementById('chartMemberCount')\n";
	echo "var myMemberCountPieChart = new Chart(ctxMemberCount, {\n";
	echo "  type: 'pie',\n";
	echo "  responsive: true,\n";
	echo "          data: {\n";

	echo "                  labels: [$strMemberCountLabels],\n";
	echo "                  datasets: [\n";
	echo "                  {\n";
	echo "                          fill: true,\n";
	echo "                          data: [$strMemberCountValues],\n";
	echo "                          backgroundColor: [\n";
						printRainbowRGBA(1);
	echo "                          ],\n";
	echo "                          borderColor: [\n";
						printRainbowRGBA(0.8);
	echo "                          ],\n";
	echo "                          borderWidth: 1\n";
	echo "                  }],\n";
	echo "           },\n";
	echo "		options: {\n";
	printPieChartOptions();
	echo "          },\n";
	echo "});\n";


	// Message aging pie chart
	// Message age: 0, 1-7, 8-14, 15-30, 31-60, 61-180, 181-365, >365
	$strMemberCountLabels = "'Current Day','2-7 Days','8-14 Days','15-30 Days', '31-60 Days','61-180 Days','181-365 Days','>365 Days'";
	$strMemberCountValues = "{$arrayMsgAging['0']}, {$arrayMsgAging['1']}, {$arrayMsgAging['8']}, {$arrayMsgAging['15']}, {$arrayMsgAging['31']}, {$arrayMsgAging['61']}, {$arrayMsgAging['181']}, {$arrayMsgAging['366']}";

	echo "var ctxMsgAging= document.getElementById('chartMessageAging')\n";
	echo "var myMsgAgingPieChart = new Chart(ctxMsgAging, {\n";
	echo "  type: 'pie',\n";
	echo "  responsive: true,\n";
	echo "          data: {\n";

	echo "                  labels: [$strMemberCountLabels],\n";
	echo "                  datasets: [\n";
	echo "                  {\n";
	echo "                          fill: true,\n";
	echo "                          data: [$strMemberCountValues],\n";
	echo "                          backgroundColor: [\n";
	echo "                                  'rgba(180, 29, 20, 1)',\n";
        echo "                                  'rgba(188, 110, 20, 1)',\n";
        echo "                                  'rgba(227, 234, 7, 1)',\n";
        echo "                                  'rgba(151, 234, 7, 1)',\n";
        echo "                                  'rgba(16, 170, 8, 1)',\n";
        echo "                                  'rgba(8, 149, 170, 1)',\n";
        echo "                                  'rgba(8, 13, 170, 1)',\n";
        echo "                                  'rgba(132, 8, 170, 1)'\n";
	echo "                          ],\n";
	echo "                          borderColor: [\n";
	echo "                                  'rgba(180, 29, 20, 0.8)',\n";
        echo "                                  'rgba(188, 110, 20, 0.8)',\n";
        echo "                                  'rgba(227, 234, 7, 0.8)',\n";
        echo "                                  'rgba(151, 234, 7, 0.8)',\n";
        echo "                                  'rgba(16, 170, 8, 0.8)',\n";
        echo "                                  'rgba(8, 149, 170, 0.8)',\n";
        echo "                                  'rgba(8, 13, 170, 0.8)',\n";
        echo "                                  'rgba(132, 8, 170, 0.8)'\n";
	echo "                          ],\n";
	echo "                          borderWidth: 1\n";
	echo "                  }],\n";
	echo "           },\n";
	echo "		options: {\n";
	printPieChartOptions();
	echo "          },\n";
	echo "});\n";

	$strDateLabels = NULL;
	$strPrivateMessages = NULL;
	$strChannelMessages = NULL;
	$strCalls = NULL;
	$strMeetings = NULL;
	$strSVTTeams = NULL;
	$strSVTSkype = NULL;
	$iAllPrivateMessages = 0;
	$iAllChannelMessages = 0;
	$iAllCalls = 0;
	$iAllMeetings = 0;	

	$strFilterStart = (new \DateTime('UTC'))->format('Y-01-01\T00:00:00\Z');
	$strFilterEnd = (new \DateTime('UTC'))->format('Y-12-31\T23:59:59\Z');

	$listMsgCounts = $web->getLists()->getByTitle("TeamsAndSkypeUsage");
	$itemsMsgCounts = $listMsgCounts->getItems()->filter("ReportDate gt '{$strFilterStart}' and ReportDate le '{$strFilterEnd}'")->top(375);

	$ctxClientConnection->load($itemsMsgCounts);
	$ctxClientConnection->executeQuery();

	foreach( $itemsMsgCounts->getData() as $dataMsgCount ) {
		$iS4BChatMessages = $dataMsgCount->SkypeIM;
		$iPrivateMessages = $dataMsgCount->TeamsPrivateChats;
		$iChannelMessages = $dataMsgCount->TeamsTeamConversations;
		$iCalls = $dataMsgCount->TeamsCalls;
		$iMeetings = $dataMsgCount->TeamsMeetings;

		$iPercentTeamsMsgs = $iPrivateMessages / ($iPrivateMessages + $iS4BChatMessages);

		$arrayDateLabel = explode("T",$dataMsgCount->ReportDate);
		$dateReportDate = $arrayDateLabel[0];
		$iAllPrivateMessages = $iAllPrivateMessages + $iPrivateMessages;
		$iAllChannelMessages = $iAllChannelMessages + $iChannelMessages;
		$iAllCalls = $iAllCalls + $iCalls;
		$iAllMeetings = $iAllMeetings + $iMeetings;
		if($strDateLabels){
			$strDateLabels = "{$strDateLabels}, \"$dateReportDate\"";
			$strPrivateMessages = $strPrivateMessages . ", " . $iPrivateMessages;
			$strChannelMessages = $strChannelMessages . ", " . $iChannelMessages;
			$strCalls = $strCalls . ", " . $iCalls;
			$strMeetings = $strMeetings . ", " . $iMeetings;
			$strSVTTeams = $strSVTTeams . ", " . $iPercentTeamsMsgs;
			$strSVTSkype = $strSVTSkype . ", " . (1 - $iPercentTeamsMsgs);
		}
		else{
			
			$strDateLabels = "\"$dateReportDate\"";
			$strPrivateMessages = $iPrivateMessages;
			$strChannelMessages = $iChannelMessages;
			$strCalls = $iCalls;
			$strMeetings = $iMeetings;
			$strSVTTeams = $iPercentTeamsMsgs;
			$strSVTSkype = (1 - $iPercentTeamsMsgs);
		}
	}

	// YTD Messages By Type
	echo "var ctxMsgCounts = document.getElementById('chartMessageCounts')\n";
	echo "var myMsgCountPieChart = new Chart(ctxMsgCounts, {\n";
	echo "  type: 'pie',\n";
	echo "  responsive: true,\n";
	echo "          data: {\n";
	echo "                  labels: [\"Private Messages\", \"Team Chat Messages\", \"Calls\", \"Meetings\"],\n";
	echo "                  datasets: [\n";
	echo "                  {\n";
	echo "                          fill: true,\n";
	echo "                          data: [$iAllPrivateMessages, $iAllChannelMessages, $iAllCalls, $iAllMeetings],\n";
	echo "                          backgroundColor: [\n";
						printRainbowRGBA(1);
	echo "                          ],\n";
	echo "                          borderColor: [\n";
						printRainbowRGBA(0.8);
	echo "                          ],\n";
	echo "                          borderWidth: 1\n";
	echo "                  }],\n";
	echo "           },\n";
	echo "          options: {\n";
	printPieChartOptions();
	echo "          },\n";
	echo "});\n";

	// Line chart private chat count history
	echo "var ctxPrivateChatHistory= document.getElementById('chartPrivateChats')\n";
	echo "var myPrivateChatHistoryChart = new Chart(ctxPrivateChatHistory, {\n";
	echo "  type: 'line',\n";
	echo "  responsive: true,\n";
	echo "          data: {\n";

	echo "                  labels: [$strDateLabels],\n";
	echo "                  datasets: [\n";
	echo "                  {\n";
	echo "                          fill: false,\n";
	echo "                          data: [$strPrivateMessages],\n";
	echo "                          backgroundColor: [\n";
	echo "                                  'rgba(99, 88, 255, 0.8)',\n";
	echo "                          ],\n";
	echo "                          borderColor: [\n";
	echo "                                  'rgba(99, 88, 255, 1)',\n";
	echo "                          ],\n";
	echo "                          borderWidth: 3\n";
	echo "                  }],\n";
	echo "           },\n";
        echo "          options: {\n";
        printLineChartOptions();
        echo "          }\n";
	echo "});\n";

	echo "var ctxChannelChatHistory= document.getElementById('chartTeamChats')\n";
	echo "var myChannelChatHistoryChart = new Chart(ctxChannelChatHistory, {\n";
	echo "  type: 'line',\n";
	echo "  responsive: true,\n";
	echo "          data: {\n";

	echo "                  labels: [$strDateLabels],\n";
	echo "                  datasets: [\n";
	echo "                  {\n";
	echo "                          fill: false,\n";
	echo "                          data: [$strChannelMessages],\n";
	echo "                          backgroundColor: [\n";
	echo "                                  'rgba(99, 88, 255, 0.8)',\n";
	echo "                          ],\n";
	echo "                          borderColor: [\n";
	echo "                                  'rgba(99, 88, 255, 1)',\n";
	echo "                          ],\n";
	echo "                          borderWidth: 3\n";
	echo "                  }],\n";
	echo "           },\n";
	echo "		options: {\n";
        printLineChartOptions();
        echo "          }\n";
	echo "	});\n";

	// Line chart channel chat count history
	echo "var ctxCallsHistory= document.getElementById('chartCalls')\n";
	echo "var myCallsHistoryChart = new Chart(ctxCallsHistory, {\n";
	echo "  type: 'line',\n";
	echo "  responsive: true,\n";
	echo "          data: {\n";

	echo "                  labels: [$strDateLabels],\n";
	echo "                  datasets: [\n";
	echo "                  {\n";
	echo "                          fill: false,\n";
	echo "                          data: [$strCalls],\n";
	echo "                          backgroundColor: [\n";
	echo "                                  'rgba(99, 88, 255, 0.8)',\n";
	echo "                          ],\n";
	echo "                          borderColor: [\n";
	echo "                                  'rgba(99, 88, 255, 1)',\n";
	echo "                          ],\n";
	echo "                          borderWidth: 3\n";
	echo "                  }],\n";
	echo "           },\n";
	echo "		options: {\n";
        printLineChartOptions();
        echo "          }\n";
	echo "	});\n";

	// Line chart channel chat count history
	echo "var ctxMeetingsHistory= document.getElementById('chartMeetings')\n";
	echo "var myMeetingsHistoryChart = new Chart(ctxMeetingsHistory, {\n";
	echo "  type: 'line',\n";
	echo "  responsive: true,\n";
	echo "          data: {\n";

	echo "                  labels: [$strDateLabels],\n";
	echo "                  datasets: [\n";
	echo "                  {\n";
	echo "                          fill: false,\n";
	echo "                          data: [$strMeetings],\n";
	echo "                          backgroundColor: [\n";
	echo "                                  'rgba(99, 88, 255, 0.8)',\n";
	echo "                          ],\n";
	echo "                          borderColor: [\n";
	echo "                                  'rgba(99, 88, 255, 1)',\n";
	echo "                          ],\n";
	echo "                          borderWidth: 3\n";
	echo "                  }],\n";
	echo "           },\n";
	echo "		options: {\n";
        printLineChartOptions();
        echo "          }\n";
	echo "	});\n";

	$strUpgradeDateLabels = NULL;
	$arrayUpgradeBuckets = array();

	$listUpgradeCounts = $web->getLists()->getByTitle("Teams Interoperability");
	$itemsUpgradeCounts = $listUpgradeCounts->getItems()->filter("Interoperability_x0020_Level eq 'UpgradeToTeams'")->top(20000);

	$ctxClientConnection->load($itemsUpgradeCounts);
	$ctxClientConnection->executeQuery();

	$iUpgradedToTeams = 0;
	foreach( $itemsUpgradeCounts->getData() as $dataUpgradeCount ) {
		$iUpgradedToTeams = $iUpgradedToTeams + 1;
		$arrayUpgradeDateLabel = explode("T",$dataUpgradeCount->Created);
		$arrayUpgradeBuckets[$arrayUpgradeDateLabel[0]] = $arrayUpgradeBuckets[$arrayUpgradeDateLabel[0]] + 1;
	}
	ksort($arrayUpgradeBuckets);
	$strCreationTimeLabels = NULL;
	$strCreationTimeValues = NULL;
	foreach($arrayUpgradeBuckets as $strUpgradeBucket=>$iUpgradeValue){
		//$strUpgradeBucket = '"' . mb_substr($strUpgradeBucket, 0, 4) . "-" . mb_substr($strUpgradeBucket, 4, 2) . "-" . mb_substr($strUpgradeBucket, 6, 2) . '"';
		if($strCreationTimeLabels){
		$strCreationTimeLabels = "{$strCreationTimeLabels}, \"$strUpgradeBucket\"";
            $strCreationTimeValues = $strCreationTimeValues . ", " . $iUpgradeValue;
		}
		else{
			$strCreationTimeLabels = "\"$strUpgradeBucket\"";
			$strCreationTimeValues = $iUpgradeValue;
		}
	}
	
	// line chart of people upgrading to Teams
	echo "var ctxUpgradesHistory= document.getElementById('chartUpgrades')\n";
	echo "var myUpgradesHistoryChart = new Chart(ctxUpgradesHistory, {\n";
	echo "  type: 'line',\n";
	echo "  responsive: true,\n";
	echo "          data: {\n";

	echo "                  labels: [$strCreationTimeLabels],\n";
	echo "                  datasets: [\n";
	echo "                  {\n";
	echo "                          fill: false,\n";
	echo "                          data: [$strCreationTimeValues],\n";
	echo "                          backgroundColor: [\n";
	echo "                                  'rgba(99, 88, 255, 0.8)',\n";
	echo "                          ],\n";
	echo "                          borderColor: [\n";
	echo "                                  'rgba(99, 88, 255, 1)',\n";
	echo "                          ],\n";
	echo "                          borderWidth: 3\n";
	echo "                  }],\n";
	echo "           },\n";
	echo "		options: {\n";
	echo "			title: {\n";
	echo "				display: true,\n";
	echo "				fontColor: 'white',\n";
	echo "				fontStyling: 'normal',\n";
	echo "				fontSize: '14',\n";
	echo "				text: \"$iUpgradedToTeams accounts have been upgraded to Teams\"\n,";
	echo "			},\n";
        printLineChartOptions();
        echo "          }\n";
	echo "	});\n";

        // Line chart channel chat count history
        echo "var ctxSVTHistory= document.getElementById('chartSkypeVSTeams')\n";
        echo "var mySVTHistoryChart = new Chart(ctxSVTHistory, {\n";
        echo "  type: 'line',\n";
        echo "  responsive: true,\n";
        echo "          data: {\n";

        echo "                  labels: [$strDateLabels],\n";
        echo "                  datasets: [\n";
        echo "                  {\n";
        echo "                          fill: true,\n";
        echo "                          data: [$strSVTSkype],\n";
        echo "                          backgroundColor: [\n";
        echo "                                  'rgba(0, 162, 232, 1)',\n";
        echo "                          ],\n";
        echo "                          borderColor: [\n";
        echo "                                  'rgba(0, 162, 232, 0.8)',\n";
        echo "                          ],\n";
        echo "                          borderWidth: 3\n";
        echo "                  },\n";
        echo "			{\n";
        echo "                          fill: true,\n";
        echo "                          data: [$strSVTTeams],\n";
        echo "                          backgroundColor: [\n";
        echo "                                  'rgba(83, 89, 175, 1)',\n";
        echo "                          ],\n";
        echo "                          borderColor: [\n";
        echo "                                  'rgba(83, 89, 175, 0.8)',\n";
        echo "                          ],\n";
        echo "                          borderWidth: 3\n";
        echo "                  },\n";
        echo "			],\n";
        echo "           },\n";
        echo "          options: {\n";
	echo "                  defaultFontColor: 'white',\n";
	echo "                  legend: {\n";
	echo "                          display: false,\n";
	echo "                          labels: {\n";
	echo "                                  fontColor: 'white'\n";
	echo "                          }\n";
	echo "                  },\n";
	echo "                  elements: {\n";
	echo "                          point: { hitRadius: 5, hoverRadius: 5, radius: 0, pointStyle: 'circle', backgroundColor: 'rgba(255,255,87,0.8)', } \n";
	echo "                  },\n";
	echo "                  tooltips: {\n";
	echo "                          mode: 'point'\n";
	echo "                  },\n";
	echo "                  multiTooltipTemplate : \"<%=datasetLabel%> : <%=value%>\",\n";
	echo "                  events: ['mousemove', 'mouseout', 'click', 'touchstart', 'touchmove'],\n";
	echo "                  scales: {\n";
	echo "                          yAxes: [{\n";
	echo "					stacked: true,\n";
	echo "                                  gridLines: {\n";
	echo "                                          display: true,\n";
	echo "                                          drawBorder: true,\n";
	echo "                                          drawOnChartArea: true,\n";
	echo "                                          zeroLineColor: 'white',\n";
	echo "                                          color: 'rgba(255,255,255,0.7)',\n";
	echo "                                          fontColor: 'white'\n";
	echo "                                  },\n";
	echo "                                  ticks: {\n";
	echo "						min: 0,\n";
	echo "						max: 1,\n";
	echo "                                          callback: function(value, index, values) {\n";
	echo "                                                  return (value * 100) + '%';\n";
	echo "                                          },\n";
	echo "                                          beginAtZero: true,\n";
	echo "                                          fontColor: 'white',\n";
	echo "                                  }\n";
	echo "                          }],\n";
	echo "                          xAxes: [{\n";
	echo "                                  gridLines: {\n";
	echo "                                          display: true,\n";
	echo "                                          drawBorder: true,\n";
	echo "                                          drawOnChartArea: false,\n";
	echo "                                          zeroLineColor: 'white',\n";
	echo "                                          color: 'rgba(255,255,255,0.7)',\n";
	echo "                                          fontColor: 'white'\n";
	echo "                                  },\n";
	echo "                                  ticks: {\n";
	echo "                                          fontColor: 'white',\n";
	echo "                                          autoSkip: true,\n";
	echo "                                          maxTicksLimit: 20\n";
	echo "                                  }\n";
	echo "                          }]\n";
	echo "                  }\n";
        echo "          }\n";
        echo "  });\n";

		
?>

function drawChart(evt, chartName) {
  var i, tabcontent, tablinks;

  tabcontent = document.getElementsByClassName("tabcontent");
  for (i = 0; i < tabcontent.length; i++) {
    tabcontent[i].style.display = "none";
  }

  tablinks = document.getElementsByClassName("tablinks");
  for (i = 0; i < tablinks.length; i++) {
    tablinks[i].className = tablinks[i].className.replace(" active", "");
  }

  document.getElementById(chartName).style.display = "block";
  evt.currentTarget.className += " active";
} 

</script>

</body>
