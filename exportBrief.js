console.log('exportBrief.js is connected');
// TODO: add logo, format Date/Time data, remove gutter between the table row columns, style the word doc

// added this function because sp.js wouldn't load
function Getdata() {
  try {
    clientContext = new SP.ClientContext.get_current();
  } catch (err) {
    alert(err);
  }
}

// TODO: Store brief ID in a variable ->  uncomment this code when push on prod

// const briefNum = (() => {
//   // look for the URL parameter/value
//   let query = window.location.search.substring(1);
//   // create array with param and value
//   let numArr = query.split('=');
//   // returns brief ID
//   return numArr[1];
// })()


// Convert 'False' into 'No', 'True' into 'Yes'
function convertToYesNo(value) {
  (value == false) ? value = 'No': value = 'Yes'
  return value
}

// Convert to Date time friendly display
function dateFriendly(d) {
  return String.format('{0:dd}-{0:MM}-{0:yyyy} {0:HH}:{0:mm}:{0:ss}', new Date(d));
}

// ============== get the item from the BRIEFS list ==============
var targetListItem;

// returns the context information about the current web application
siteUrl = 'https://bauer.sharepoint.com/sites/Sandbox/UK-Radio-CreationHub';
clientContext = new SP.ClientContext(siteUrl);
targetList = clientContext.get_web().get_lists().getByTitle('Briefs'); // our Briefs list

// WHEN USER CLICKS EXPORT button ==> Retrieve item from BRIEFS list with the ID number
function retrieveBriefInfo(briefNum) {
  targetListItem = targetList.getItemById(briefNum); // TODO: when push on prod REMOVE THIS param and replace with briefNum
  clientContext.load(targetListItem);
  clientContext.executeQueryAsync(Function.createDelegate(this, this.retrOnQuerySucceeded), Function.createDelegate(this, this.retrOnQueryFailed));
}

// IF query succeds
function retrOnQuerySucceeded() {
  // all fields for the brief
  const briefName = targetListItem.get_item('BriefName');
  const submitBy = targetListItem.get_item('Author').$7W_1;
  const CRMID = targetListItem.get_item('CRMOpportunityID');
  let bauerCreate = convertToYesNo(targetListItem.get_item('BauerCreate'));

  const client = targetListItem.get_item('Client');
  const clientContact = targetListItem.get_item('ClientContact');
  const webAddress = targetListItem.get_item('WebsiteAddress');
  const deadlineToAM = dateFriendly(targetListItem.get_item('DeadlineToAM'));
  const tagLine = targetListItem.get_item('ExistingTaglineOrCreative');
  let stations = '';

  let stationsVals = targetListItem.get_item('StationsToRunActivityOn');
  for (let i = 0; i < stationsVals.length; i++) {
    stations += `${stationsVals[i].get_lookupValue()} , `;
  }

  const campaignTarget = targetListItem.get_item('WhoCampaignTargetting');
  const call2Action = targetListItem.get_item('CallToAction');
  const whatsIn = targetListItem.get_item('WhatsInForListener');
  const insights = targetListItem.get_item('Insights');
  const campaignLength = targetListItem.get_item('TotalLengthCampaign');
  const budget = targetListItem.get_item('CampaignBudget');
  const goLive = dateFriendly(targetListItem.get_item('DateToGoAirGoLive'));


  const planitNumber = targetListItem.get_item('PlanitNumber');
  const CHRef = targetListItem.get_item('CHRef');
  const streamsReq = targetListItem.get_item('StreamsRequired');

  // Fields for CREATIVE stream
  const otherInfo = targetListItem.get_item('OtherUsefulInfo');
  const ammunition = targetListItem.get_item('HelpfulAmmunition');
  const singleMindProp = targetListItem.get_item('SingleMindedProposition');
  const coopNum = targetListItem.get_item('CoopNumber');

  // Fields for DIGITAL
  const endDate = dateFriendly(targetListItem.get_item('EndDate'));
  const DGProductsReq = targetListItem.get_item('DigitalPackageRequired');
  const typeRespReq = targetListItem.get_item('TypeOfResponseRequired');
  const notes = targetListItem.get_item('Notes');

  // Fields for AUDIO
  const prodType = targetListItem.get_item('ProductionType');
  const reqType = targetListItem.get_item('RequestType');
  const aboutBrief = targetListItem.get_item('AdditionalInformation');
  const scriptIdeas = targetListItem.get_item('ScriptIdeas');
  const sendFinalCopyTo = targetListItem.get_item('DeliveryInstructions');
  const scriptReqDate = dateFriendly(targetListItem.get_item('ScriptDueDate'));
  const audioReqDate = dateFriendly(targetListItem.get_item('AudioRequiredDate'));
  const transDate = targetListItem.get_item('TxDate');
  const duration = targetListItem.get_item('Duration');
  const confidential = convertToYesNo(targetListItem.get_item('Confidential'));

  // Fields for S&P
  const suggestions = targetListItem.get_item('CampaignSuggestions');

  // Fields for SS
  const commercialType = targetListItem.get_item('CommercialType');

  // set the HTML structure
  bodyHTML = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word'><head><meta charset='utf-8'><title>Export Brief to Word doc</title></head><body style='font-family: Calibri'><div><div style='text-align: center;align-items: center;'><h1> CREATION HUB </h1><br> <table>";

  const fieldsObj = {
    'Brief name': briefName,
    'Submitted by': submitBy,
    'CRM opportunity ID': CRMID,
    'Bauer Create': bauerCreate,
    'Client': client,
    'Client contact': clientContact,
    'Website address': webAddress,
    'Deadline to AM': deadlineToAM,
    'Stations to run activity on': stations,
    'Existing tagline': tagLine,
    'Campaign targeting': campaignTarget,
    'Call to action': call2Action,
    "What's in it for the listener?": whatsIn,
    'Client insights': insights,
    'Total length of campaign': campaignLength,
    'Total Campaign budget (£)': budget,
    'Date to go live': goLive,
    'PlanIT number': planitNumber,
    'streams required': streamsReq
  };

  // push different streams related extra key/values to the fields object
  // for Creative stream
  if (streamsReq.includes('Creative')) {
    fieldsObj['Other useful information'] = otherInfo;
    fieldsObj['Helpful ammunition'] = ammunition;
    fieldsObj['Single minded proposition'] = singleMindProp;
    fieldsObj['COOP number'] = coopNum;
  }
  // for Digital stream
  if (streamsReq.includes('Digital')) {
    fieldsObj['End Date'] = endDate;
    fieldsObj['Digital products required'] = DGProductsReq;
    fieldsObj['Type of response required'] = typeRespReq;
    fieldsObj['Notes'] = notes;
  }
  // for Audio Production (SI) stream
  if (streamsReq.includes('Audio Production')) {
    fieldsObj['Production type'] = prodType;
    fieldsObj['Request type'] = reqType;
    fieldsObj['Tell us about your brief...'] = aboutBrief;
    fieldsObj['Put any script ideas here'] = scriptIdeas;
    fieldsObj['Who do you want us to send the final copy to...'] = sendFinalCopyTo;
    fieldsObj['Script Required Date'] = scriptReqDate;
    fieldsObj['Audio Required Date'] = audioReqDate;
    fieldsObj['Transmission Date'] = transDate;
    fieldsObj['Duration'] = duration;
    fieldsObj['Confidential'] = confidential;
  }

  // for S&P stream
  if (streamsReq.includes('S and P')) {
    fieldsObj['Campaign suggestions & revelant additional information'] = suggestions;
  }
  // for SS stream
  if (streamsReq.includes('Sales Support')) {
    fieldsObj['Commercial type'] = commercialType;
  }

  // Loop all fields key/value
  for (let key in fieldsObj) {
    // conditional to sort them by Streams category
    if (key.includes('Other useful information')) {
      bodyHTML += `<tr style='border: 1px solid #3777ff; background-color: #EA9E8D; padding: 20px'><h2> CREATIVE</h2></tr>`;
    }
    if (key.includes('End Date')) {
      bodyHTML += `<tr style='border: 1px solid #3777ff; background-color: #677db7; padding: 20px'><h2> DIGITAL </h2></tr>`;
    }
    if (key.includes('Production type')) {
      bodyHTML += `<tr style='border: 1px solid #3777ff; background-color: #B5D6D6; padding: 20px'><h2> AUDIO PRODUCTION </h2></tr>`;
    }
    if (key.includes('Campaign suggestions & revelant additional information')) {
      bodyHTML += `<tr style='border: 1px solid #3777ff; background-color: #FAC05E; padding: 20px'><h2> S&P </h2></tr>`;
    }
    if (key.includes('Commercial type')) {
      bodyHTML += `<tr style='border: 1px solid #3777ff; background-color: #AF7595; padding: 20px'><h2> SALES SUPPORT </h2></tr>`;
    }
    if (fieldsObj[key] === null) {
      fieldsObj[key] = ''
    };

    bodyHTML += `<tr style='border: 1px solid #3777ff; padding: 20px'>
        <td style='border: 1px solid #3777ff; width: 80px; padding: 20px'> ${key}</td>
        <td style='border: 1px solid #3777ff; padding: 20px'> ${fieldsObj[key]} </td>
        </tr>`;
  }

  bodyHTML += '</table></body></html>';

  console.log(bodyHTML);

  Export2Word(bodyHTML, CHRef);
}

function onQueryFailed(sender, args) {
  alert('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

// Will create the word document
function Export2Word(bodyHTML, CHRef) {
  var blob = new Blob(['\ufeff', bodyHTML], {
    type: 'application/msword'
  });

  // Specify link url
  var url = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(bodyHTML);

  // Create download link element
  var downloadLink = document.createElement('a');

  document.body.appendChild(downloadLink);

  if (navigator.msSaveOrOpenBlob) {
    navigator.msSaveOrOpenBlob(blob, filename);
  } else {
    // Create a link to the file
    downloadLink.href = url;

    // Setting the file name
    downloadLink.download = `${CHRef}.doc`;

    //triggering the function
    downloadLink.click();
  }

  document.body.removeChild(downloadLink);
}