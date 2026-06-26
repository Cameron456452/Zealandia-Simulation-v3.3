// Global variables
let ss = SpreadsheetApp.getActiveSpreadsheet();

// Creates menu
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts') // Add .addItem() calls to add to the list
      .addItem('Automation', 'automationSheet')
      .addItem('Automation – Copying', 'automationSheetCopy')
      .addItem('Automation – Incrementing', 'automationSheetInc')
      .addItem('Automation – Appending', 'automationSheetApp')
      .addItem('Goods', 'goods')
      .addItem('Statistics', 'historyReports')
      .addItem('Population', 'settlePopulation')
      .addItem('Savings', 'settleSavings')
      .addItem('Election', 'conductElection')
      .addToUi();
}

let automationSS = ss.getSheetByName("Automation");
function automationSheet() {
  // Goods
  goods();
  // Copying
  let copyingNumber = automationSS.getRange("E1").getValue();
  let copyFromData = automationSS.getRange(2, 1, copyingNumber, 4).getValues();
  for (let i = 0; i < copyingNumber; i++) {
    let copyFromRangeName = copyFromData[i][0];
    let copyFromSheetName = copyFromData[i][1];
    let copyToRangeName = copyFromData[i][2];
    let copyToSheetName = copyFromData[i][3];
    ss.getSheetByName(copyFromSheetName).getRange(copyFromRangeName).copyTo(ss.getSheetByName(copyToSheetName).getRange(copyToRangeName), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  }
  // Incrementing
  let incrementingNumber = automationSS.getRange("J1").getValue();
  let incrementData = automationSS.getRange(2, 7, incrementingNumber, 3).getValues();
  for (let i = 0; i < incrementingNumber; i++) {
    let incrementRangeName = incrementData[i][0];
    let incrementSheetName = incrementData[i][1];
    let incrementAmount = incrementData[i][2];
    ss.getSheetByName(incrementSheetName).getRange(incrementRangeName).setValue(ss.getSheetByName(incrementSheetName).getRange(incrementRangeName).getValue() + incrementAmount);
  }
  // Appending
  let automationNumber = automationSS.getRange("Q1").getValue();
  let appendingData = automationSS.getRange(2, 12, automationNumber, 5).getValues();
  for (let i = 0; i < automationNumber; i++) {
    let appendFromRangeName = appendingData[i][0];
    let appendFromSheetName = appendingData[i][1];
    let appendToRowNumber = appendingData[i][2];
    let appendToColumnNumber = appendingData[i][3];
    let appendToSheetName = appendingData[i][4];
    let appendToSheet = ss.getSheetByName(appendToSheetName);
    for (let j = appendToRowNumber; j <= appendToRowNumber + 100; j++) {
      if (appendToSheet.getRange(j, appendToColumnNumber).getValue() == "") {
        ss.getSheetByName(appendFromSheetName).getRange(appendFromRangeName).copyTo(appendToSheet.getRange(j, appendToColumnNumber), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
        break;
      }
    }
  }
}
function automationSheetCopy() {
  // Copying
  let copyingNumber = automationSS.getRange("E1").getValue();
  let copyFromData = automationSS.getRange(2, 1, copyingNumber, 5).getValues();
  for (let i = 0; i < copyingNumber; i++) {
    let copyFromRangeName = copyFromData[i][0];
    let copyFromSheetName = copyFromData[i][1];
    let copyToRangeName = copyFromData[i][2];
    let copyToSheetName = copyFromData[i][3];
    ss.getSheetByName(copyFromSheetName).getRange(copyFromRangeName).copyTo(ss.getSheetByName(copyToSheetName).getRange(copyToRangeName), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  }
}
function automationSheetInc() {
  // Incrementing
  let incrementingNumber = automationSS.getRange("J1").getValue();
  let incrementData = automationSS.getRange(2, 7, incrementingNumber, 3).getValues();
  for (let i = 0; i < incrementingNumber; i++) {
    let incrementRangeName = incrementData[i][0];
    let incrementSheetName = incrementData[i][1];
    let incrementAmount = incrementData[i][2];
    ss.getSheetByName(incrementSheetName).getRange(incrementRangeName).setValue(ss.getSheetByName(incrementSheetName).getRange(incrementRangeName).getValue() + incrementAmount);
  }
}
function automationSheetApp() {
  // Appending
  let automationNumber = automationSS.getRange("Q1").getValue();
  let appendingData = automationSS.getRange(2, 12, automationNumber, 5).getValues();
  for (let i = 0; i < automationNumber; i++) {
    let appendFromRangeName = appendingData[i][0];
    let appendFromSheetName = appendingData[i][1];
    let appendToRowNumber = appendingData[i][2];
    let appendToColumnNumber = appendingData[i][3];
    let appendToSheetName = appendingData[i][4];
    let appendToSheet = ss.getSheetByName(appendToSheetName);
    for (let j = appendToRowNumber; j <= appendToRowNumber + 100; j++) {
      if (appendToSheet.getRange(j, appendToColumnNumber).getValue() == "") {
        ss.getSheetByName(appendFromSheetName).getRange(appendFromRangeName).copyTo(appendToSheet.getRange(j, appendToColumnNumber), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
        break;
      }
    }
  }
}
function goods() {
  let goodsSheet = ss.getSheetByName("Goods");
  let goodAmount = goodsSheet.getRange("A1").getValue();
  let goodValues = goodsSheet.getRange("A2:M15").getValues();
  // Update new prices
  console.log(goodValues);
  for (let i = 2; i <= goodAmount; i++) {
    console.log(i);
    goodsSheet.getRange(i, 3).setValue(goodValues[i-2][2] * (1 + goodValues[i-2][12]));
  }
  // Move people around
  for (let i = 2; i <= goodAmount; i++) {
    // Badly-performing goods lose workers, good ones gain them
    console.log(goodsSheet.getRange(i, 7));
    console.log((1 + (goodsSheet.getRange(i, 4).getValue()/2.5)));

    goodsSheet.getRange(i, 7).setValue(goodValues[i-2][6] * (1 + (goodsSheet.getRange(i, 4).getValue()/2.5)));

    console.log(goodsSheet.getRange(i, 7));
  }
  // Normalize worker numbers
  goodsSheet.getRange("G2:G15").setValues(goodsSheet.getRange("S2:S15").getValues());

  // Update new prices for new economy
  for (let i = 2; i <= goodAmount; i++) {
    goodsSheet.getRange(i, 3).setValue(goodValues[i-2][2] * (1 + goodValues[i-2][12]));
  }

  // Update consumption
  consumption();
}

function consumption() {
  let consumeSheet = ss.getSheetByName("Cost of living");
  let consumeTable = consumeSheet.getRange("B5:BE18").getValues();
  let wealthMins = consumeSheet.getRange("B4:BE4").getValues()[0];
  let n = 14;
  let totalGoods = [Array(n).fill(0), Array(n).fill(0), Array(n).fill(0), Array(n).fill(0)]

  let vult = consumeSheet.getRange("D21:F27").getValues();
  let tas = consumeSheet.getRange("J21:L37").getValues();
  let cook = consumeSheet.getRange("P21:R37").getValues();
  let south = consumeSheet.getRange("V21:X37").getValues();
  let allProvinces = [vult, tas, cook, south]



  for(let province = 0; province < allProvinces.length; province++) {
    for(let income = 0; income < allProvinces[province].length; income++) {
      // find wealth
      let wealth = allProvinces[province][income][1];
      // adds goods to array
      for(let goods = 0; goods < totalGoods[province].length; goods++) {
        totalGoods[province][goods] += ((consumeTable[goods][wealth]) * (allProvinces[province][income][2]) * (allProvinces[province][income][0] / wealthMins[wealth]));
      }
    }
  }

  const totalGoodsColumnArrays = totalGoods.map(arr => arr.map(val => [val]));

  consumeSheet.getRange("B40:B53").setValues(totalGoodsColumnArrays[0]);
  consumeSheet.getRange("I40:I53").setValues(totalGoodsColumnArrays[1]);
  consumeSheet.getRange("O40:O53").setValues(totalGoodsColumnArrays[2]);
  consumeSheet.getRange("U40:U53").setValues(totalGoodsColumnArrays[3]);
  
}

function settlePopulation() {
  let populationSheet = ss.getSheetByName("Vulturalia");
  for (let i = 0; i < 3; i++) {
    populationSheet.getRange('G46:H146').setValues(populationSheet.getRange('I46:J146').getValues());
    populationSheet.getRange('N52').setValue(populationSheet.getRange('N52').getValue());
  }
  populationSheet = ss.getSheetByName("Tasmania");
  for (let i = 0; i < 3; i++) {
    populationSheet.getRange('G46:H146').setValues(populationSheet.getRange('I46:J146').getValues());
    populationSheet.getRange('N52').setValue(populationSheet.getRange('N52').getValue());
  }
  populationSheet = ss.getSheetByName("Cooksland");
  for (let i = 0; i < 3; i++) {
    populationSheet.getRange('G46:H146').setValues(populationSheet.getRange('I46:J146').getValues());
    populationSheet.getRange('N52').setValue(populationSheet.getRange('N52').getValue());
  }
  populationSheet = ss.getSheetByName("Southland");
  for (let i = 0; i < 3; i++) {
    populationSheet.getRange('G46:H146').setValues(populationSheet.getRange('I46:J146').getValues());
    populationSheet.getRange('N52').setValue(populationSheet.getRange('N52').getValue());
  }
}
function settleSavings() {
  let savingsSheet = ss.getSheetByName("Vulturalia");
  for (let i = 0; i < 5; i++) {
    savingsSheet.getRange('P16:P32').setValues(savingsSheet.getRange('Q16:Q32').getValues());
    savingsSheet.getRange('R16:R32').setValues(savingsSheet.getRange('S16:S32').getValues());
  }
  savingsSheet = ss.getSheetByName("Tasmania");
  for (let i = 0; i < 5; i++) {
    savingsSheet.getRange('P16:P32').setValues(savingsSheet.getRange('Q16:Q32').getValues());
    savingsSheet.getRange('R16:R32').setValues(savingsSheet.getRange('S16:S32').getValues());
  }
  savingsSheet = ss.getSheetByName("Cooksland");
  for (let i = 0; i < 5; i++) {
    savingsSheet.getRange('P16:P32').setValues(savingsSheet.getRange('Q16:Q32').getValues());
    savingsSheet.getRange('R16:R32').setValues(savingsSheet.getRange('S16:S32').getValues());
  }
  savingsSheet = ss.getSheetByName("Southland");
  for (let i = 0; i < 5; i++) {
    savingsSheet.getRange('P16:P32').setValues(savingsSheet.getRange('Q16:Q32').getValues());
    savingsSheet.getRange('R16:R32').setValues(savingsSheet.getRange('S16:S32').getValues());
  }
}

// Stalin stuff
let partyNumber = 13;
let provinces = 4;
let voters = [];
let votes = new Array(partyNumber).fill(0);
let provincalVotes = [Array(partyNumber).fill(0), Array(partyNumber).fill(0), Array(partyNumber).fill(0), Array(partyNumber).fill(0), 
Array(partyNumber).fill(0)];
let provBallotsCasted = new Array(provinces).fill(0);

let simVotingSheet = ss.getSheetByName("Demographics");
let partisanship = simVotingSheet.getRange("S5").getValue();

let election = ss.getSheetByName("SimElection");

let demographicNames = ["Gender", "Age", "Income", "Politics", "Location", "Province", "Ethnicity"];
let demographicWeights = [1, 1.5, 2.5, partisanship, 1.5, 0.5, 2];

// let humanLikelihoodData = [
        
//         // gender
//         [simVotingSheet.getRange("F23").getValue(), simVotingSheet.getRange("F24").getValue()], 

//         [simVotingSheet.getRange("F48").getValue(), simVotingSheet.getRange("F49").getValue(), simVotingSheet.getRange("F50").getValue(), 
//         simVotingSheet.getRange("F51").getValue()], 

//         [simVotingSheet.getRange("F27").getValue(), simVotingSheet.getRange("F28").getValue(), simVotingSheet.getRange("F29").getValue(),
//         simVotingSheet.getRange("F30").getValue(), simVotingSheet.getRange("F31").getValue()], 

//         [simVotingSheet.getRange("F34").getValue(), simVotingSheet.getRange("F35").getValue(), simVotingSheet.getRange("F36").getValue(),
//         simVotingSheet.getRange("F37").getValue(), simVotingSheet.getRange("F38").getValue()], 
        
//         [simVotingSheet.getRange("F41").getValue(), simVotingSheet.getRange("F42").getValue()],

//         [simVotingSheet.getRange("B45").getValue(), simVotingSheet.getRange("C45").getValue(), simVotingSheet.getRange("D45").getValue(),
//         simVotingSheet.getRange("E45").getValue(),],

//         [simVotingSheet.getRange("F2").getValue(), simVotingSheet.getRange("F3").getValue(), simVotingSheet.getRange("F4").getValue(), 
//         simVotingSheet.getRange("F5").getValue(),],

//         ];

// gender
let values = simVotingSheet.getRange("B23:E24").getValues();
let gender = values[0].map((_, colIndex) =>
  values.map(row => row[colIndex])
);

// age
let values2 = simVotingSheet.getRange("B48:E51").getValues();
let age = values2[0].map((_, colIndex) =>
  values2.map(row => row[colIndex])
);

// income 
let values3 = simVotingSheet.getRange("B27:E31").getValues();
let incomeList = values3[0].map((_, colIndex) =>
  values3.map(row => row[colIndex])
);

// politics
let values4 = simVotingSheet.getRange("B34:E38").getValues();
let politics = values4[0].map((_, colIndex) =>
  values4.map(row => row[colIndex])
);

// Location
let values5 = simVotingSheet.getRange("B41:E42").getValues();
let location = values5[0].map((_, colIndex) =>
  values5.map(row => row[colIndex])
);

// Province
// (Has to be determined first)
// let values = simVotingSheet.getRange("B34:E38").getValues();
// let province = values[0].map((_, colIndex) =>
//   values.map(row => row[colIndex])
// );

let provinceList = [
  [simVotingSheet.getRange("B45").getValue(), simVotingSheet.getRange("C45").getValue(), simVotingSheet.getRange("D45").getValue(),
        simVotingSheet.getRange("E45").getValue()],
  [simVotingSheet.getRange("B45").getValue(), simVotingSheet.getRange("C45").getValue(), simVotingSheet.getRange("D45").getValue(),
        simVotingSheet.getRange("E45").getValue()],
  [simVotingSheet.getRange("B45").getValue(), simVotingSheet.getRange("C45").getValue(), simVotingSheet.getRange("D45").getValue(),
        simVotingSheet.getRange("E45").getValue()],
  [simVotingSheet.getRange("B45").getValue(), simVotingSheet.getRange("C45").getValue(), simVotingSheet.getRange("D45").getValue(),
        simVotingSheet.getRange("E45").getValue()],
]

// Ethnicity
let values6 = simVotingSheet.getRange("B2:E6").getValues();
let ethnicity = values6[0].map((_, colIndex) =>
  values6.map(row => row[colIndex])
);

humanLikelihoodData = [gender, age, incomeList, politics, location, provinceList, ethnicity]

let voterCategoryNames = election.getRange("A2:A40").getValues();
let partyPopularities = election.getRange("B2:N40").getValues();
let demographicVotes = Array.from(Array(humanLikelihoodData.length-1), () => new Array(partyNumber).fill(0));
// let weightIndices = [[0,0],[0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0],[0,0,0,0],[0,0,0,0]];
// let weightIndices = [
//                      [[0,0],[0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0],[0,0,0,0],[0,0,0,0]],
//                      [[0,0],[0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0],[0,0,0,0],[0,0,0,0]],
//                      [[0,0],[0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0],[0,0,0,0],[0,0,0,0]],
//                      [[0,0],[0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0],[0,0,0,0],[0,0,0,0]],
//                     ]

// console.log(humanLikelihoodData);

// for (let i = 0; i < humanLikelihoodData.length; i++) {
//   for (let k = 0; k < humanLikelihoodData[i].length; k++) {
//     for (let j = 0; j < humanLikelihoodData[i][k].length; j++) {

//       console.log(humanLikelihoodData[i][k][j]);

//       if (j !== 0) {
//         weightIndices[i][k][j] =
//           humanLikelihoodData[i][k][j] + weightIndices[i][k][j - 1];
//       } else {
//         weightIndices[i][k][j] = humanLikelihoodData[i][k][j];
//       }

//       if (j === humanLikelihoodData[i][k].length - 1) {
//         weightIndices[i][k][j] = 1;
//       }
//     }
//   }
// }

// console.log(weightIndices[0]);

let weightIndices = [];

for (let i = 0; i < humanLikelihoodData.length; i++) {
  weightIndices[i] = [];
  for (let k = 0; k < humanLikelihoodData[i].length; k++) {
    weightIndices[i][k] = [];
    let cumulative = 0;
    for (let j = 0; j < humanLikelihoodData[i][k].length; j++) {
      cumulative += humanLikelihoodData[i][k][j];
      // force the very last index to be exactly 1
      if (j === humanLikelihoodData[i][k].length - 1) {
        cumulative = 1;
      }
      weightIndices[i][k][j] = cumulative;
    }
  }
}

function conductElection() {
  // let voters = Math.round(election.getRange("M8").getValues() / 100);
  let voters = election.getRange("M62").getValues();

  const categories = {
    0: ["Male", "Female"],
    1: ["18-29", "30-49", "50-64", "65+"],
    2: ["Poor", "Lower Middle", "Middle", "Upper Middle", "Upper"],
    3: ["Far left", "Center left", "Centrist", "Center right", "Far right"],
    4: ["Urban", "Rural"],
    5: ["Makaurau", "Tasmania", "Cooksland", "Southland"],
    6: ["Anglo", "Dutch", "Maori", "Asian", "Zealandian"]
  };

  let voterData = [];

  // Build a dictionary for each category
  for (let index in categories) {
    voterData[index] = {};
    categories[index].forEach(option => {
      voterData[index][option] = new Array(partyNumber).fill(0)
    });
  }

  
  for(let i=0; i < voters; i++)
  {

    if(i % 10000 == 0) {
      console.log(i);
    }

    Person = new Voter();
    ballot = Person.createVoter();
    votes[ballot[0]] += 1;
    if(Person.getIdentity(5) == "Makaurau")
    {
      provincalVotes[0][ballot[1]] += 1;
      provBallotsCasted[0] += 1;
    }

    if(Person.getIdentity(5) == "Tasmania")
    {
      provincalVotes[1][ballot[1]] += 1;
      provBallotsCasted[1] += 1;
    }

    if(Person.getIdentity(5) == "Cooksland")
    {
      provincalVotes[2][ballot[1]] += 1;
      provBallotsCasted[2] += 1;
    }

    if(Person.getIdentity(5) == "Southland")
    {
      provincalVotes[3][ballot[1]] += 1;
      provBallotsCasted[3] += 1;
    }

    // Add to exit poll
    // Categories: 0–6
    for(let j=0; j < 7; j++) {
      voterData[j][Person.getIdentity(j)][ballot[0]] += 1;
    }

  }
  for (let k = 0; k < votes.length; k++) {
    election.getRange(42, k+2).setValue(votes[k]/voters);
    election.getRange(43, k+2).setValue(provincalVotes[0][k]/provBallotsCasted[0]);
    election.getRange(44, k+2).setValue(provincalVotes[1][k]/provBallotsCasted[1]);
    election.getRange(45, k+2).setValue(provincalVotes[2][k]/provBallotsCasted[2]);
    election.getRange(46, k+2).setValue(provincalVotes[3][k]/provBallotsCasted[3]);
  }

  let rowsDown = -2;

  for (let categoryIndex = 0; categoryIndex < voterData.length; categoryIndex++) {
    let category = voterData[categoryIndex];

    // console.log("=== Category", categoryIndex, "===");
    rowsDown += 1;
    for (let [label, numbers] of Object.entries(category)) {
      // console.log("Label:", label);
      
      rowsDown += 1;
      for (let i = 0; i < numbers.length; i++) {
        // console.log("   Index:", i, "Value:", numbers[i]);
        election.getRange(87+rowsDown, i+2).setValue(numbers[i]/numbers.reduce((acc, val) => acc + val, 0));
      }
    }
  }

  console.log(voterData);

}

class Voter {
  
  constructor()
  {
    this.province = provinceChoice(5, ["Makaurau", "Tasmania", "Cooksland", "Southland"]);
    this.gender = randChoices(0, ["Male", "Female"], provinceNumber(this.province));
    this.age = randChoices(1, ["18-29", "30-49", "50-64", "65+"], provinceNumber(this.province));
    this.income = randChoices(2, ["Poor", "Lower Middle", "Middle", "Upper Middle", "Upper"], provinceNumber(this.province));
    this.politics = randChoices(3, ["Far left", "Center left", "Centrist", "Center right", "Far right"], provinceNumber(this.province));
    this.location = randChoices(4, ["Urban", "Rural"], provinceNumber(this.province))
    this.ethnicity = randChoices(6, ["Anglo", "Dutch", "Maori", "Asian", "Zealandian"], provinceNumber(this.province));
    
    this.identity = [this.gender, this.age, this.income, this.politics, this.location, this.province, this.ethnicity];
  }

  getIdentity(index)
  {
    return this.identity[index];
  }
  
  createVoter()
  {
    // Raw data for approvals by demographic group and population
    // [[Party], [Party]]
    // Inner []: [Asian Approval for Party A, Lower Middle Class Approval for Party A]
    let voterGroup = this.createHuman();
    let voterDesire = new Array(partyNumber).fill(0); // If error, add more 0s (needs to be as many as the party)
    let provinceDesire = [];

    for(let i=0; i < voterGroup.length; i++) // parties
    {
      for(let j=0; j < voterGroup[i].length; j++) // demographic groups except provinces
      {
        // tiny bit of randomness to prevent ties
        voterDesire[i] += ratingScore(voterGroup[i][j]) * demographicWeights[j] + (Math.random()*0.0001); 
      }

    }

    for(let i=0; i < voterGroup.length; i++)
    {
      provinceDesire.push(ratingScore(voterGroup[i][5]) * 2 + voterDesire[i]);
    }

    let max = voterDesire[0];
    let provinceMax = provinceDesire[0];
    let index = 0;
    let provinceIndex = 0;
    for(let i=1; i < voterDesire.length; i++)
    {
      if(voterDesire[i] > max)
      {
        max = voterDesire[i];
        index = i;
      }
      else if(voterDesire[i] == max) // gives first party an advantage due to ties
      {
        if(Math.random() > 0.5)
        {
          max = voterDesire[i];
          index = i;
          console.log("Tie")
        }
      }
    }

    for(let i=1; i < provinceDesire.length; i++)
    {
      if(provinceDesire[i] > provinceMax)
      {
        provinceMax = provinceDesire[i];
        provinceIndex = i;
      }
      else if(provinceDesire[i] == provinceMax) // gives first party an advantage due to ties
      {
        if(Math.random() > 0.5)
        {
          provinceMax = provinceDesire[i];
          provinceIndex = i;
        }
      }
    }

  // Makes splt ticket voting rarer, 50% chance of voting by province, 40% chance of voting by national, 10% chance of split
    // if (index != provinceIndex) {
    //   let flip = Math.random();
    //   if (flip < 0.5) {
    //     index = provinceIndex;
    //   } else if (flip > 0.6) {
    //     provinceIndex = index;
    //   }
    // }

    return [index, provinceIndex]; // Index of the party in party list
  }

  createHuman()
  {
    let voterList = new Array(partyNumber).fill(null).map(() => []); // If there's an error, add more boxes until equal to partyNumber

    for(let i=0; i < partyNumber; i++)
    {
      for(let j=0; j < this.identity.length; j++)
      {
        voterList[i].push(addVoterList(this.identity[j], i)); // Gets each party's popularity in each demographic
      }

    }

    return voterList;
  }
}

function ratingScore(approval) {
  let score = 0
  let chance = Math.random()

  if(chance+(1) <= approval/100) // die-hard supporter 
  {
    score = 125
  }
  else if(chance+(2/3) <= approval/100) // die-hard supporter 
  {
    score = Math.random()*(125-100) + 100; // 100 to 125
  }
  else if(chance+(1/3) <= approval/100) // fanatic supporter 
  {
    score = 100; // 100
  }
  else if(chance <= approval/100) // supporter
  {
    score = Math.random()*(75-50) + 50; // 50 to 75
  }
  else if(chance-(1/3) <= approval/100) // critic
  {
    score = Math.random()*(50-25) + 25; // 25 to 50
  }
  else if(chance-(2/3) <= approval/100) // hater
  {
    score = Math.random()*(25-0) + 0; // 0 to 25
  }
  else if(chance-(1) <= approval/100) // die-hard hater
  {
    score = Math.random()*(50-25) - 25; // -25 to 0
  }
  else { // hater
    score = -25; // 0
  }

  return Math.round(score);
}

function addVoterList(variable, partyColumn)
{

  for(let i=0; i <= partyPopularities.length - 1; i++)
  {
    
    if(voterCategoryNames.length == 1)
    {
      voterCategoryNames[i] = voterCategoryNames[i][0];
    }

    if(voterCategoryNames[i] == variable) // Pre-stores the range into a 2D list for faster computations
    {
      return partyPopularities[i][partyColumn];
    }
  }
}

function provinceNumber(prov) {
    if(prov == "Makaurau")
    {
      return 0;
    }

    if(prov == "Tasmania")
    {
      return 1;
    }

    if(prov == "Cooksland")
    {
      return 2;
    }

    if(prov == "Southland")
    {
      return 3;
    }

    return 99;
}

function randChoices(indicesIndex, choiceList, provIdx)
{
  let option = Math.random();

  for(let i=0; i < weightIndices[indicesIndex][provIdx].length; i++)
  {
    if(weightIndices[indicesIndex][provIdx][i] >= option) {
      return choiceList[i];
    }
  }
  return choiceList[weightIndices[indicesIndex].length-1]; // Just in case, return the last one
}

function provinceChoice(indicesIndex, choiceList) {
  let option = Math.random();

  for(let i=0; i < weightIndices[indicesIndex][0].length; i++)
  {
    if(weightIndices[indicesIndex][0][i] >= option) {
      return choiceList[i];
    }
  }
  return choiceList[weightIndices[indicesIndex].length-1]; // Just in case, return the last one
}

// Decay program
function decay() {
  let partyPopularities = election.getRange("B2:N40").getValues();

  for(let i=0; i < partyPopularities.length; i++)  {
    for(let j = 0; j < partyPopularities[i].length; j++) {
      if(partyPopularities[i][j] > 50) {
        partyPopularities[i][j] -= Math.round(partyPopularities[i][j] * 0.1);

        if(partyPopularities[i][j] < 50) {
          partyPopularities[i][j] = 50;
        }
      }
    }
  }

  for (let j = 0; j < partyPopularities.length; j++) {
    election
    .getRange(
      2,
      2,
      partyPopularities.length,   
      partyPopularities[0].length      
    )
    .setValues(partyPopularities);
  }

  return 0;
}

function indexOfMax(arr) { // Borrowed this code from the Overflow
  if (arr.length === 0) {
    return -1;
  }
  let max = arr[0];
  let maxIndex = 0;
  for (let i = 1; i < arr.length; i++) {
    if (arr[i] > max || (arr[i] == max && Math.random() >= 0.5)) { // Break ties with a coinflip
      maxIndex = i;
      max = arr[i];
    }
  }
  return maxIndex;
}
function sumArr(arr) { // This is faster than reduce by about two times
  let sum = 0;
  for (let z = 0; z < arr.length; z++) {
    if (!isNaN(z)) { // Ignore all non-numbers
      sum += Number(arr[z]);
    }
  }
  return sum;
}