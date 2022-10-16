// @ts-nocheck
var id_arkusz1 = "*****";

var Arkusz1 = SpreadsheetApp.openById(id_arkusz1);
var Arkusz_wynik1 = SpreadsheetApp.getActiveSpreadsheet();

function zliczanie(){

22
// -- Wyliczenie ile kolumn jest wypełnionych - referencje -- //

 var ilosc_kolumn1 = Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getRange(7,10,1,Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getLastColumn()).getDisplayValues();
 var ostatnia_kolumna1 = Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getLastColumn();

 var ilosc_k1 = [];

 for(var i=0; i<ostatnia_kolumna1; i++){
    if(ilosc_kolumn1[i] !=""){
      ilosc_k1.push(ilosc_kolumn1[0][i]);
    }
 }

// --- Stworzenie siatki wartosci i kolorów --- //

 var zakres_stator1 =  Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getRange(8,10, Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getLastRow(), ilosc_k1.length).getDisplayValues();
 var zakres_stator2 =  Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getRange(8,2, Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getLastRow(), 1).getDisplayValues();
 //var kolor1 = Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getRange(8,2, Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getLastRow(), 1).getBackgrounds();
 var znak1 = Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getRange(8,1, Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getLastRow(), 1).getValues();
 var zakres_opisu1 = Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getRange(8,2, Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getLastRow(),1).getDisplayValues();

// --- Zliczanie wyniku --- //



var wynik1 = [];
var referencja = [];
var opis = [];


for(var x=0; x<ilosc_k1.length; x++){
        var ref1 = ilosc_kolumn1[0][x];
              for(var y=0; y<zakres_stator1.length; y++){
                       if(znak1[y] == "x"){
                           var opis1 = zakres_stator2[y][0];
                           var wartosc1 = zakres_stator1[y-2][x]; 
                           var wartosc2=zakres_stator1[298][x];
                           
                              if(zakres_opisu1[y] == opis1  && (wartosc1 !== "NOK") && wartosc2==''){
                                 wynik1.push(wartosc1);
                                 referencja.push(ref1);
                                 opis.push(opis1);
                                 
                              }
                       } 
                      
              }
}




Call = zapis_w_arkuszu(wynik1, referencja, opis);

}

function zapis_w_arkuszu(wynik1, referencja, opis){

 // --- stworzenie listy referencji - usunięcie duplikatów --- //

   var lista_referencji = [];
   var proces = [];

   for(var x=0; x<referencja.length; x++){
    //if(referencja[x] != referencja[x+1] && opis[x] != opis[x+1]){
     // lista_referencji.push('\n');
      lista_referencji.push(referencja[x]);
      proces.push(opis[x]);
   // }
   }
   
   // --- Pobranie zdefiniowanej listy i podliczenie OK --- //

   var referencja_sprawdzona = [];

   for(var j=0; j<lista_referencji.length; j++){
     var jaka_ref = lista_referencji[j];
     var opis_procesu = proces[j]; 
     var ile_OK = 0;
      for(var k=0; k<referencja.length; k++){
        if(referencja[k] == jaka_ref && proces[k] == opis_procesu){
            ile_OK = ile_OK +1;
        }
      }
      referencja_sprawdzona.push([opis_procesu, jaka_ref, ile_OK]);
   }

   // --- wpisanie wyników -- //

   for(var i=0; i<referencja_sprawdzona.length; i++){
    Arkusz_wynik1.getSheetByName("xMG STATOR proto2 DVPV parts").appendRow([referencja_sprawdzona[i][0], referencja_sprawdzona[i][1], referencja_sprawdzona[i][2]]);

   }

  
}

function zliczanie_v2(){

// -- Wyliczenie ile kolumn jest wypełnionych - referencje -- //

 var ilosc_kolumn1 = Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getRange(7,10,1,Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getLastColumn()).getDisplayValues();
 var ostatnia_kolumna1 = Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getLastColumn();

 var ilosc_k1 = [];

 for(var i=0; i<ostatnia_kolumna1; i++){
    if(ilosc_kolumn1[0][i] != ""){
      ilosc_k1.push(ilosc_kolumn1[0][i]);
    }
 }

 // -- Pobranie zawartości zakładki -- //

var zakres = Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getRange(13,1,Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getLastRow(), Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getLastColumn()).getDisplayValues();

// -- Zdefiniowanie czego szukamy -- //

var szukam = []; // " STATOR STACK", "SLOT INSULATION", "INSULATION COINING "

for(var i=0; i<zakres.length; i++){
   if(zakres[i][0] === "x"){
      szukam.push(zakres[i][1]);
   }
}


// --- Proces przeszukiwania --- //

var lista = [];

for(var x=0; x<ilosc_k1.length; x++){ // kolumny
   for(var y=0; y<zakres.length; y++){ // linijki
     for( var i=0; i<szukam.length; i++){ // szukamy
          if(zakres[y][1] === szukam[i] && zakres[297][x] != "OK"){
            var proces = szukam[i-1]; //zakres[y][1];
            var weryfikacja = zakres[y-2][x+9];
            var referencja = zakres[1][x+9]; //6
                             lista.push([proces, referencja, weryfikacja]);           
          }    
     }
   }
}



// -- Usunięcie starych danych -- //

var wyniki = Arkusz_wynik1.getSheetByName("xMG STATOR proto2 DVPV parts").getRange(2,1, Arkusz_wynik1.getSheetByName("xMG STATOR proto2 DVPV parts").getLastRow(),8).clearContent();

// -- Wpislanie wyniku w zakładkę -- //

for(var z=0; z<lista.length; z++){
  if(lista[z][2] === "" && lista[z-1][2] === "OK"){
   Arkusz_wynik1.getSheetByName("xMG STATOR proto2 DVPV parts").appendRow([lista[z][0],lista[z][1],1]);
  // Arkusz_wynik1.getSheetByName("Monitoring ONLINE - UPDATED").getRange(1,1,50,50).setValue([lista[z][0]]);
  }
}

}




// Rotor xMG 25 kW proto2






// @ts-nocheck
var id_arkusz2 = "*****";

var Arkusz2 = SpreadsheetApp.openById(id_arkusz2);
var Arkusz_wynik2 = SpreadsheetApp.getActiveSpreadsheet();

function zliczanie1(){


// -- Wyliczenie ile kolumn jest wypełnionych - referencje -- //

 var ilosc_kolumn2 = Arkusz2.getSheetByName("xMG25kW - PROTO2").getRange(3,9,1,Arkusz2.getSheetByName("xMG25kW - PROTO2").getLastColumn()).getDisplayValues();
 var ostatnia_kolumna2 = Arkusz2.getSheetByName("xMG25kW - PROTO2").getLastColumn();

 var ilosc_k2 = [];

 for(var i=0; i<ostatnia_kolumna2; i++){
    if(ilosc_kolumn2[i] !=""){
      ilosc_k2.push(ilosc_kolumn2[0][i]);
    }
 }

// --- Stworzenie siatki wartosci i kolorów --- //

 var zakres_rotor1 =  Arkusz2.getSheetByName("xMG25kW - PROTO2 - CAL257").getRange(6,9, Arkusz2.getSheetByName("xMG25kW - PROTO2").getLastRow(), ilosc_k2.length).getDisplayValues();
 var zakres_rotor2 =  Arkusz2.getSheetByName("xMG25kW - PROTO2 - CAL257").getRange(6,2, Arkusz2.getSheetByName("xMG25kW - PROTO2").getLastRow(), 1).getDisplayValues();
 //var kolor1 = Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getRange(8,2, Arkusz1.getSheetByName("Stator PROTO2 (DVPV PARTS)").getLastRow(), 1).getBackgrounds();
 var znak2 = Arkusz2.getSheetByName("xMG25kW - PROTO2 - CAL257").getRange(6,1, Arkusz2.getSheetByName("xMG25kW - PROTO2").getLastRow(), 1).getValues();
 var zakres_opisu2 = Arkusz2.getSheetByName("xMG25kW - PROTO2 - CAL257").getRange(6,2, Arkusz2.getSheetByName("xMG25kW - PROTO2").getLastRow(),1).getDisplayValues();

// --- Zliczanie wyniku --- //



var wynik2 = [];
var referencja1 = [];
var opis2 = [];

// tutaj skończyłem

for(var x=0; x<ilosc_k2.length; x++){
        var ref2 = ilosc_kolumn2[0][x];
              for(var y=0; y<zakres_rotor1.length; y++){
                       if(znak2[y].includes("x")){
                           var opis3 = zakres_rotor2[y][0];
                           var wartosc3 = zakres_rotor1[y-2][x]; 
                           var wartosc4=zakres_rotor1[229][x];
                           
                              if(zakres_opisu2[y] == opis3  && (wartosc3 !== "NOK") && wartosc4==''){
                                 wynik2.push(wartosc3);
                                 referencja1.push(ref2);
                                 opis2.push(opis3);
                                 
                              }
                       } 
                      
              }
}




Call = zapis_w_arkuszu1(wynik2, referencja1, opis2);

}

function zapis_w_arkuszu1(wynik2, referencja1, opis2){

 // --- stworzenie listy referencji - usunięcie duplikatów --- //

   var lista_referencji1 = [];
   var proces1 = [];

   for(var x=0; x<referencja1.length; x++){
    //if(referencja[x] != referencja[x+1] && opis[x] != opis[x+1]){
     // lista_referencji.push('\n');
      lista_referencji1.push(referencja1[x]);
      proces1.push(opis2[x]);
   // }
   }
   
   // --- Pobranie zdefiniowanej listy i podliczenie OK --- //

   var referencja_sprawdzona1 = [];

   for(var j=0; j<lista_referencji1.length; j++){
     var jaka_ref1 = lista_referencji1[j];
     var opis_procesu1 = proces1[j]; 
     var ile_OK1 = 0;
      for(var k=0; k<referencja1.length; k++){
        if(referencja1[k] == jaka_ref1 && proces1[k] == opis_procesu1){
            ile_OK1 = ile_OK1 +1;
        }
      }
      referencja_sprawdzona1.push([opis_procesu1, jaka_ref1, ile_OK1]);
   }

   // --- wpisanie wyników -- //

   for(var i=0; i<referencja_sprawdzona1.length; i++){
    Arkusz_wynik2.getSheetByName("xMG ROTOR proto2").appendRow([referencja_sprawdzona1[i][0], referencja_sprawdzona1[i][1], referencja_sprawdzona1[i][2]]);

   }

  
}

function zliczanie_v21(){

// -- Wyliczenie ile kolumn jest wypełnionych - referencje -- //

 var ilosc_kolumn2 = Arkusz2.getSheetByName("xMG25kW - PROTO2").getRange(3,9,1,Arkusz2.getSheetByName("xMG25kW - PROTO2").getLastColumn()).getDisplayValues();
 var ostatnia_kolumna2 = Arkusz2.getSheetByName("xMG25kW - PROTO2").getLastColumn();

 var ilosc_k2 = [];

 for(var i=0; i<ostatnia_kolumna2; i++){
    if(ilosc_kolumn2[0][i] != ""){
      ilosc_k2.push(ilosc_kolumn2[0][i]);
    }
 }

 // -- Pobranie zawartości zakładki -- //

var zakres1 = Arkusz2.getSheetByName("xMG25kW - PROTO2").getRange(8,1,Arkusz2.getSheetByName("xMG25kW - PROTO2 - CAL257").getLastRow(), Arkusz2.getSheetByName("xMG25kW - PROTO2").getLastColumn()).getDisplayValues();

// -- Zdefiniowanie czego szukamy -- //

var szukam1 = []; // " Process1", "Process2", "Process2 "

for(var i=0; i<zakres1.length; i++){
   if(zakres1[i][0].includes("x")){
      szukam1.push(zakres1[i][1]);
   }
}


// --- Proces przeszukiwania --- //

var lista1 = [];

for(var x=0; x<ilosc_k2.length; x++){ // kolumny
   for(var y=0; y<zakres1.length; y++){ // linijki
     for( var i=0; i<szukam1.length; i++){ // szukamy
          if(zakres1[y][1] === szukam1[i] && zakres1[228][x] != "OK"){
            var proces1 = szukam1[i-1]; //zakres[y][1];
            var weryfikacja1 = zakres1[y-2][x+8];
            var referencja1 = zakres1[14][x+8]; //6
                             lista1.push([proces1, referencja1, weryfikacja1]);           
          }    
     }
   }
}



// -- Usunięcie starych danych -- //

var wyniki1 = Arkusz_wynik2.getSheetByName("xMG ROTOR proto2").getRange(2,1, Arkusz_wynik2.getSheetByName("xMG ROTOR proto2").getLastRow(),8).clearContent();

// -- Wpislanie wyniku w zakładkę -- //

for(var z=0; z<lista1.length; z++){
  if(lista1[z][2] === "" && lista1[z-1][2] === "OK"){
   Arkusz_wynik2.getSheetByName("xMG ROTOR proto2").appendRow([lista1[z][0],lista1[z][1],1]);
  // Arkusz_wynik1.getSheetByName("Monitoring ONLINE - UPDATED").getRange(1,1,50,50).setValue([lista[z][0]]);
  }
}

}

