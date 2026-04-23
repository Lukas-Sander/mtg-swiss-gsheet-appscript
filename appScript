//TODO: table assignment
//TODO: make match tables dynamic, aka add "buttons" to generate match tables
//TODO: make all tables no-touch and only one sheet for configuring and managing?

function onEdit(e) {
  // 1. Abbruchbedingung
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  const range = e.range;
  
  // 2. Prüfen, ob die Änderung in der richtigen Tabelle und Zelle stattfand
  if (sheet.getName() === "Standings" && range.getA1Notation() === "N1" && range.getValue() === true) {
    // WICHTIG: range mitgeben, damit wir die Checkbox später zurücksetzen können!
    calculateStandings(range);
  }

  if (sheet.getName() === "Match 1" && range.getA1Notation() === "I1" && range.getValue() === true) {
    generateMatchPairings('Match 1', range);
  }
}

function generateMatchPairings(tableName, range) {
  let pairings = [];
  let standings = getTableData('Standings');
  let assignedPlayers = new Set();

  // Gruppieren der Spieler nach Match-Punkten (Index 2)
  let groupedByScore = {};
  standings.forEach(row => {
    let score = row[2];
    if (!groupedByScore[score]) groupedByScore[score] = [];
    groupedByScore[score].push(row);
  });

  // Punktestände absteigend sortieren, um die generelle Rangfolge beizubehalten
  let sortedScores = Object.keys(groupedByScore).sort((a, b) => b - a);

  let shuffledStandings = [];

  // Jede Gruppe intern shuffeln (Fisher-Yates-Algorithmus) und zusammenfügen
  sortedScores.forEach(score => {
    let group = groupedByScore[score];
    for (let i = group.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [group[i], group[j]] = [group[j], group[i]];
    }
    shuffledStandings.push(...group);
  });

  standings = shuffledStandings;

  standings.forEach((row, index) => {
    if(row[0] !== 1 && !assignedPlayers.has(row[1])) {  //only find an opponent if player hasn't dropped, and if player hasnt been assigned an opponent yet

      let prevOpponents = row[11].split(';');
      let hasOpp = false;
    
      /**
       * find opponent with closest standing that hasnt played against them yet:
       * 1. try next standings row as opponent
       * 2. if next row is either dropped, already assigned or in prevOpponents, try the row below that
       * 3. if next row is nonexistent, player gets a bye
       * 4. if player already had a bye, we need to re-evaluate all players from the bottom up again and find one who hasnt had a bye yet. that player gets the bye and swaps with the current player
       */
      for(let i = index+1; i< standings.length; i++) {
        const opp = standings[i];
        if(opp[0] === 1) {  //opp dropped
          continue;
        }
        if(assignedPlayers.has(opp[1])) { //opp already assigned
          continue;
        }
        if(prevOpponents.includes(opp[1])) {  //opp already playes against player
          continue;
        }

        //at this point we have a match for the player
        assignedPlayers.add(opp[1]);
        assignedPlayers.add(row[1]);

        const matchRow = [
          row[1],
          opp[1],
          'tischplaceholder', //TODO: table assignment
          0,
          0,
          0,
          0
        ];

        pairings.push(matchRow);
        hasOpp = true;
        break;
      }

      if(!hasOpp) {
          //at this point we couldn't find a match for the player, so they get a bye. but we need to check if this player already had a bye
          if(prevOpponents.includes('bye')) {
            //TODO: reverse traverse pairings and match player with next possible opponent to the top. switch pairings so that player gets a bye. rinse and repeat until no double bye exists any more
          }
          else {
            assignedPlayers.add(row[1]);
            const matchRow = [
              row[1],
              'bye',
              'tischplaceholder',
              0,
              0,
              0,
              0
            ];
            pairings.push(matchRow);
          }
      }
    }

  });


    //write match pairings into table
    const pairingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tableName);
    // getRange(StartZeile, StartSpalte, AnzahlZeilen, AnzahlSpalten)
    pairingSheet.getRange(2, 1, pairings.length, pairings[0].length).setValues(pairings);

  // do not set to false after finishing, to avoid accidentally re-pairing matches
  // if (range) range.setValue(false);
}

function calculateStandings(range) {
  const players = getTableData("Teilnehmende");
  let standings = getTableData("Standings"); // let statt const, da wir es evtl. neu bauen
  
  const matchData = [
      getTableData('Match 1'),
      getTableData('Match 2'),
      getTableData('Match 3'),
      getTableData('Match 4'),
      getTableData('Match 5')
  ];

  // if standings table is empty, fill with players first
  if(standings.length === 0) {
    
    for(let i = 0; i < players.length; i++) {
      const playerName = players[i][0]; // Name aus Spalte A
      
      // Neues Array für die Zeile pushen
      standings.push([
        0,      // dropped value. 1 means dropped
        playerName, // player name (sauber als String)
        0,  // match points
        0,  // wins
        0,  // draws
        0,  // losses
        0.33, // MWP
        0,  // played rounds
        0.33, // OMWP
        0.33, // GWP
        0.33,  // OGWP
        ''  //prevOpponents
      ]);
    }
    
    const standingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Standings");
    // getRange(StartZeile, StartSpalte, AnzahlZeilen, AnzahlSpalten)
    standingsSheet.getRange(2, 1, standings.length, standings[0].length).setValues(standings);
    
  } else {
    //first, make a players object for easy referencing
    const playerDb = {};
    players.forEach(player => {
      playerDb[player[1]] = { //this is the player name, which is unique by design
        isDropped: player[0] === 1,
        matchpoints: 0,
        wins: 0,
        draws: 0,
        losses: 0,
        mpw: 0.33,
        playedRounds: 0,
        omwp: 0.33,
        gwp: 0.33,
        ogwp: 0.33,
        prevOpponents: []
      };
    });

    //first, calculate matchpoints to playedRounds
    for (const [name, data] of Object.entries(playerDb)) {
      //find matches in all match tables
    }

    // score all individual players
    //TODO: convert into object containing players with name=key. then we can put their data and match data in there without needing to re-search again and again
    players.forEach(player => {
      
      // find player row in standings table and then start calculating
      const row = standings.findIndex(row => row.includes(player));
      if(row >= 0) {
        console.log(row);
        let playerRow = standings[row];
        matchData.forEach(matchTable => {
          // find player row and col (we need to know if they are a or b) in match table and then start calculating
          let r = -1, c = -1;
          for (let rO = 0; rO < matchTable.length; rO++) {
            const c = matchTable[rO].indexOf(search);
            if (c !== -1) { r = rO; col = c; break; }
          }
          if(r !== -1 && c !== -1) {
            //found player
            const rowScore = '' //TODO: continue working here
          }


          //TODO: also for each match table, add the prevOpponent to the row column data. this is only done here
        });

      }

    });

    // TODO: after scoring all players, sort by score
  }

  // set to false after finishing
  if (range) range.setValue(false);
}

function getTableData(tableName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tableName);
  if (!sheet) return []; // Sicherheits-Check, falls Blatt nicht existiert
  
  const lastRow = sheet.getLastRow();
  
  // Verhindert Fehler, wenn die Tabelle außer der Überschrift leer ist
  if (lastRow < 2) return []; 
  
  const lastCol = sheet.getLastColumn();
  const columnData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return columnData;
}
