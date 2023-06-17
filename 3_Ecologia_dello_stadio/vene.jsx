main();

function main() {
    var pw = 360; // larghezza della pagina (in mm)
    var ph = 270; // lunghezza della pagina (in mm)

    //La prima definisce il documento inDesign su cui operare
    var doc_setup = app.activeDocument;
    var doc = doc_setup; //la variabile doc definisce il documento aperto
    

    //Questa funzione sfrutta il motore di ricerca integrato in inDesign per trovare nel documento specificato (primo argomento) le iterazione della parola specificata (secondo argomento)
    //e restituisce la posizione secondo due variabili:
        // var x = Object.horizontalOffset;
        // var y = Object.baseline;
    var result_of_search = text_find_word(doc, "nemic");


    //"toSource()" prende il risultato dell funzione "text_find_word" come stringa
    //per ogni "," la stringa viene divisa in vettori
    //unisce la stringa ad un punto con una linea
    alert(
        "Questo è il risultato della ricerca:\n"+ 
        result_of_search.toSource().split(",").join("\n")
        );
    

    var origin_x = 732.5; //Definisce la posizione sull'asse delle x della circonferenza in cui confluiranno tutte le linee
    var origin_y = 540; //Definisce la posizione sull'asse delle y della circonferenza in cui confluiranno tutte le linee

    var x_1 = 25;
    var x_2 = 385;
    var x_3 = 745;
    var x_4 = 1092.5;
    var x_5 = 1465;

    var y_1 = 12.5;
    var y_2 = 257.5;
    var y_3 = 282.5;
    var y_4 = 527.5;
    var y_5 = 552.5;
    var y_6 = 797.5;
    var y_7 = 822.5;
    var y_8 = 1067.5;
    var y_9 = 1092.5;
    var y_10 = 1337.5;

    //Imposta un loop che scorra attraverso i risultati della funzione "text_find_word" e colleghi ogni coppia di coordinate alla circonferenza
    for(var i = 0; i < result_of_search.length;i++){
        var res_x = result_of_search[i].position[0]; //Richiama la posizione sull'asse delle x del risultato
        var res_y = result_of_search[i].position[1]; //Richiama la posizione sull'asse delle y del risultato
        
        var data = {
            "anchors_pagina_1_1":[
                [res_x,res_y],
                [x_1,y_2],
                [x_2,y_1],
                [x_3,y_2],
                [x_3,y_3],
                [x_3,y_4],
                [x_3,y_5],
                [745,675],
            ],

            "anchors_pagina_1_2":[
                [res_x,res_y],
                [x_2,y_1],
                [x_3,y_2],
                [x_3,y_3],
                [x_3,y_4],
                [x_3,y_5],
                [745,675],
            ],

            "anchors_pagina_1_3":[
                [res_x,res_y],
                [x_3,y_2],
                [x_3,y_3],
                [x_3,y_4],
                [x_3,y_5],
                [745,675],
            ],

            "anchors_pagina_1_4":[
                [res_x,res_y],
                [x_4,y_1],
                [x_4,y_2],
                [x_4,y_3],
                [x_4,y_4],
                [x_4,y_5],
                [1092.5,675],
            ],

            "anchors_pagina_2_1":[
                [res_x,res_y],
                [x_2,y_4],
                [x_3,y_3],
                [x_3,y_4],
                [x_3,y_5],
                [745,675],
            ],

            "anchors_pagina_2_2":[
                [res_x,res_y],
                [x_3,y_3],
                [x_3,y_4],
                [x_3,y_5],
                [745,675],
            ],

            "anchors_pagina_2_3":[
                [res_x,res_y],
                [x_4,y_4],
                [x_4,y_5],
                [1092.5,675],
            ],

            "anchors_pagina_3_1":[
                [res_x,res_y],
                [x_2,y_5],
                [x_3,y_6],
                [745,675],
            ],

            "anchors_pagina_3_2":[
                [res_x,res_y],
                [x_3,y_6],
                [745,675],
            ],

            "anchors_pagina_3_3":[
                [res_x,res_y],
                [x_4,y_5],
                [1092.5,675],
            ],

            "anchors_pagina_4_1":[
                [res_x,res_y],
                [x_2,y_8],
                [x_3,y_7],
                [x_3,y_6],
                [745,675],
            ],

            "anchors_pagina_4_2":[
                [res_x,res_y],
                [x_3,y_7],
                [x_3,y_6],
                [745,675],
            ],

            "anchors_pagina_4_3":[
                [res_x,res_y],
                [x_4,y_8],
                [x_4,y_7],
                [x_4,y_6],
                [1092.5,675],
            ],

            "anchors_pagina_5_1":[
                [res_x,res_y],
                [x_2,y_9],
                [x_3,y_10],
                [x_3,y_9],
                [x_3,y_8],
                [x_3,y_7],
                [x_3,y_6],
                [745,675],
            ],

            "anchors_pagina_5_2":[
                [res_x,res_y],
                [x_3,y_10],
                [x_3,y_9],
                [x_3,y_8],
                [x_3,y_7],
                [x_3,y_6],
                [745,675],
            ],

            "anchors_pagina_5_3":[
                [res_x,res_y],
                [x_4,y_9],
                [x_4,y_8],
                [x_4,y_7],
                [x_4,y_6],
                [1092.5,675],
            ],

            "anchors_pagina_5_4":[
                [res_x,res_y],
                [x_5,y_10],
                [x_4,y_9],
                [x_4,y_8],
                [x_4,y_7],
                [x_4,y_6],
                [1092.5,675],
            ],
        };
    

        //Crea una linea
        var gl = doc.graphicLines.add();
            //Definisce i punti di ancoraggio della linea a seconda che la parola si trovi nella:

        
            //Pagina 1_1
            if(res_x > 0 && res_x < pw && res_y > 0 && res_y < ph) {
                for (var j in data.anchors_pagina_1_1) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_1_1[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_1_1[j];            
                    }
                }
            }

            //Pagina 1_2
            if(res_x > pw && res_x < pw * 2 && res_y > 0 && res_y < ph) {
                for (var j in data.anchors_pagina_1_2) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_1_2[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_1_2[j];            
                    }
                }
            }

            //Pagina 1_3
            if(res_x > pw * 2 && res_x < pw * 3 && res_y > 0 && res_y < ph) {
                for (var j in data.anchors_pagina_1_3) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_1_3[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_1_3[j];            
                    }
                }
            }

            //Pagina 1_4
            if(res_x > pw * 3 && res_x < pw * 4 && res_y > 0 && res_y < ph) {
                for (var j in data.anchors_pagina_1_4) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_1_4[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_1_4[j];            
                    }
                }
            }

            //Pagina 2_1
            if(res_x > pw && res_x < pw * 2 && res_y > ph && res_y < ph * 2) {
                for (var j in data.anchors_pagina_2_1) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_2_1[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_2_1[j];            
                    }
                }
            }

            //Pagina 2_2
            if(res_x > pw * 2 && res_x < pw * 3 && res_y > ph && res_y < ph * 2) {
                for (var j in data.anchors_pagina_2_2) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_2_2[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_2_2[j];            
                    }
                }
            }

            //Pagina 2_3
            if(res_x > pw * 3 && res_x < pw * 4 && res_y > ph && res_y < ph * 2) {
                for (var j in data.anchors_pagina_2_3) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_2_3[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_2_3[j];            
                    }
                }
            }

            //Pagina 3_1
            if(res_x > pw && res_x < pw * 2 && res_y > ph * 2 && res_y < ph * 3) {
                for (var j in data.anchors_pagina_3_1) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_3_1[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_3_1[j];            
                    }
                }
            }

            //Pagina 3_2
            if(res_x > pw * 2 && res_x < pw * 3 && res_y > ph * 2 && res_y < ph * 3) {
                for (var j in data.anchors_pagina_3_2) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_3_2[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_3_2[j];            
                    }
                }
            }

            //Pagina 3_3
            if(res_x > pw * 3 && res_x < pw * 4 && res_y > ph * 2 && res_y < ph * 3) {
                for (var j in data.anchors_pagina_3_3) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_3_3[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_3_3[j];            
                    }
                }
            }

            //Pagina 4_1
            if(res_x > pw && res_x < pw * 2 && res_y > ph * 3 && res_y < ph * 4) {
                for (var j in data.anchors_pagina_4_1) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_4_1[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_4_1[j];            
                    }
                }
            }

            //Pagina 4_2
            if(res_x > pw * 2 && res_x < pw * 3 && res_y > ph * 3 && res_y < ph * 4) {
                for (var j in data.anchors_pagina_4_2) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_4_2[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_4_2[j];            
                    }
                }
            }

            //Pagina 4_3
            if(res_x > pw * 3 && res_x < pw * 4 && res_y > ph * 3 && res_y < ph * 4) {
                for (var j in data.anchors_pagina_4_3) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_4_3[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_4_3[j];            
                    }
                }
            }

            //Pagina 5_1
            if(res_x > pw && res_x < pw * 2 && res_y > ph * 4 && res_y < ph * 5) {
                for (var j in data.anchors_pagina_5_1) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_5_1[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_5_1[j];            
                    }
                }
            }

            //Pagina 5_2
            if(res_x > pw * 2 && res_x < pw * 3 && res_y > ph * 4 && res_y < ph * 5) {
                for (var j in data.anchors_pagina_5_2) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_5_2[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_5_2[j];            
                    }
                }
            }

            //Pagina 5_3
            if(res_x > pw * 3 && res_x < pw * 4 && res_y > ph * 4 && res_y < ph * 5) {
                for (var j in data.anchors_pagina_5_3) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_5_3[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_5_3[j];            
                    }
                }
            }

            //Pagina 5_4
            if(res_x > pw * 4 && res_x < pw * 5 && res_y > ph * 4 && res_y < ph * 5) {
                for (var j in data.anchors_pagina_5_4) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_5_4[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_5_4[j];            
                    }
                }
            }
  };
} //fine della funzione "main"

function genRand (min, max, decimalPlaces) {
    var result = Math.random() * (max - min) + min;
    var power = Math.pow(10, decimalPlaces);
    var result = Math.floor(result * power) / power;
return result;
}










function text_find_word(the_doc, word_string){
    //Le prime due funzioni svuotano le barre di ricerca del motore di ricerca integrato in inDesign
    app.findTextPreferences = NothingEnum.nothing; //Svuota il campo "trova"
    app.changeTextPreferences = NothingEnum.nothing; //Svuota il campo "sostituisci con"
  

    var result = new Array(); //Qui viene memorizzato come array il risultato della funzione
  

    app.findTextPreferences.findWhat = word_string; //Definisce la parola da cercare sulla base dei parametri inseriti nella funzione "text_find_word" richiamata nella funzione "main"
    var ft_result = the_doc.findText(); //Esegue la funzione  "findText" erestituisce un array

    //Imposta un loop in modo che la funzione sia ripetuta tante volte quante sono le iterazione della parola ricercata nel testo
    for(var i = 0; i < ft_result.length; i++){
        //Definisce le variabili x e y corrispondenti alla posizione di ogni iterazione della parola nella pagina
        var x = ft_result[i].horizontalOffset; //"horizontalOffset" -> la x di riferimento è all'inizio della parola, "endHorizontalOffset" -> la x di riferimento è alla fine della parola
        var y = ft_result[i].baseline; //"baseline" -> la y di riferimento è all'inizio della parola, "endBaseline" -> la y di riferimento è alla fine della parola

      //crea un oggetto (la parola) definito da una posizione x e y
      result.push({"word":word_string,"position":[x,y]});
      }

      //restituisce il risultato della funzione "findText"
      return result;
      } //fine della funzione "text_find_word"