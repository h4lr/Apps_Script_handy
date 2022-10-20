#Duplicate rows based on column

function hal92() {
  const ss = SpreadsheetApp.getActive();                                                          //select current spreadsheet
  const sh = ss.getSheetByName('Hoja 4');                                                       //select sheet0
  let vs = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();            //get values from the table and store it in "vs"
  sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();                //cleans the table
  let vA = [];                                                                            //define VA as an empty table
  vs.forEach((r, i) => {
    if(r.length > 185) {
      let t = r.slice(0,185);
      let u = r.slice(184);
      u.forEach((v,j) => {
        let x = r.slice(0,185);
        x[184] = v;
        vA.push(x);
      });
    } else {
      vA.push(r.slice());
    }
  });
  sh.getRange(2, 1, vA.length, vA[0].length).setValues(vA);

}
