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
    var result_of_search = text_find_word(doc, "squadr");


    //"toSource()" prende il risultato dell funzione "text_find_word" come stringa
    //per ogni "," la stringa viene divisa in vettori
    //unisce la stringa ad un punto con una linea
    alert(
        "Questo è il risultato della ricerca:\n"+ 
        result_of_search.toSource().split(",").join("\n")
        );
    
    var x_1 = 12.5;
    var x_2 = 372.5;
    var x_3 = 732.5;
    var x_4 = 1092.5;
    var x_5 = 1452.5;

    var y_1 = 30;
    var y_2 = ph + 30;
    var y_3 = ph * 2 + 30;
    var y_4 = ph * 3 + 30;

    var randUpper = genRand(7,11,0);
    var randLower = genRand(1,5,0)

    //Imposta un loop che scorra attraverso i risultati della funzione "text_find_word" e colleghi ogni coppia di coordinate alla circonferenza
    for(var i = 0; i < result_of_search.length;i++){
        var res_x = result_of_search[i].position[0]; //Richiama la posizione sull'asse delle x del risultato
        var res_y = result_of_search[i].position[1]; //Richiama la posizione sull'asse delle y del risultato
        
        var data = {
            "anchors_pagina_1_1_upper":[
                [res_x,res_y],
                [x_2,y_1 + randUpper],
                [x_3,y_1 + randLower],
                [x_3,y_2 + randUpper],
                [x_2,y_2 + randLower],
                [372.5,540],
            ],

            "anchors_pagina_1_1_lower":[
                [res_x,res_y],
                [x_2,y_1 + randLower],
                [x_3,y_1 + randUpper],
                [x_3,y_2 + randLower],
                [x_2,y_2 + randUpper],
                [372.5,540],
            ],

            "anchors_pagina_1_2_upper":[
                [res_x,res_y],
                [x_3,y_1 + randUpper],
                [x_3,y_2 + randLower],
                [x_2,y_2 + randUpper],
                [372.5,540],
            ],

            "anchors_pagina_1_2_lower":[
                [res_x,res_y],
                [x_3,y_1 + randLower],
                [x_3,y_2 + randUpper],
                [x_2,y_2 + randLower],
                [372.5,540],
            ],

            "anchors_pagina_1_3_upper":[
                [res_x,res_y],
                [x_4,y_1 + randUpper],
                [x_3,y_1 + randLower],
                [x_3,y_2 + randUpper],
                [x_2,y_2 + randLower],
                [372.5,540],
            ],

            "anchors_pagina_1_3_lower":[
                [res_x,res_y],
                [x_4,y_1 + randLower],
                [x_3,y_1 + randUpper],
                [x_3,y_2 + randLower],
                [x_2,y_2 + randUpper],
                [372.5,540],
            ],

            "anchors_pagina_2_1_upper":[
                [res_x,res_y],
                [x_1,y_2 + randUpper],
                [x_2,y_2 + randLower],
                [372.5,540],
            ],

            "anchors_pagina_2_1_lower":[
                [res_x,res_y],
                [x_1,y_2 + randLower],
                [x_2,y_2 + randUpper],
                [372.5,540],
            ],

            "anchors_pagina_2_2_upper":[
                [res_x,res_y],
                [x_2,y_2 + randUpper],
                [372.5,540],
            ],

            "anchors_pagina_2_2_lower":[
                [res_x,res_y],
                [x_2,y_2 + randLower],
                [372.5,540],
            ],

            "anchors_pagina_2_3_upper":[
                [res_x,res_y],
                [x_3,y_2 + randUpper],
                [x_2,y_2 + randLower],
                [372.5,540],
            ],

            "anchors_pagina_2_3_lower":[
                [res_x,res_y],
                [x_3,y_2 + randLower],
                [x_2,y_2 + randUpper],
                [372.5,540],
            ],

            "anchors_pagina_2_4_upper":[
                [res_x,res_y],
                [x_4,y_2 + randUpper],
                [x_3,y_2 + randLower],
                [x_2,y_2 + randUpper],
                [372.5,540],
            ],

            "anchors_pagina_2_4_lower":[
                [res_x,res_y],
                [x_4,y_2 + randLower],
                [x_3,y_2 + randUpper],
                [x_2,y_2 + randLower],
                [372.5,540],
            ],

            "anchors_pagina_3_1_upper":[
                [res_x,res_y],
                [x_2,y_3 + randUpper],
                [372.5,540],
            ],
            "anchors_pagina_3_1_lower":[
                [res_x,res_y],
                [x_2,y_3 + randLower],
                [372.5,540],
            ],

            "anchors_pagina_3_2_upper":[
                [res_x,res_y],
                [x_3,y_3 + randUpper],
                [x_2,y_3 + randLower],
                [372.5,540],
            ],

            "anchors_pagina_3_2_lower":[
                [res_x,res_y],
                [x_3,y_3 + randLower],
                [x_2,y_3 + randUpper],
                [372.5,540],
            ],

            "anchors_pagina_3_3_upper":[
                [res_x,res_y],
                [x_4,y_3 + randUpper],
                [x_3,y_3 + randLower],
                [x_2,y_3 + randUpper],
                [372.5,540],
            ],

            "anchors_pagina_3_3_lower":[
                [res_x,res_y],
                [x_4,y_3 + randLower],
                [x_3,y_3 + randUpper],
                [x_2,y_3 + randLower],
                [372.5,540],
            ],

            "anchors_pagina_4_1_upper":[
                [res_x,res_y],
                [x_2,y_4 + randUpper],
                [x_3,y_4 + randLower],
                [x_3,y_3 + randUpper],
                [x_2,y_3 + randLower],
                [372.5,540],
            ],

            "anchors_pagina_4_1_lower":[
                [res_x,res_y],
                [x_2,y_4 + randLower],
                [x_3,y_4 + randUpper],
                [x_3,y_3 + randLower],
                [x_2,y_3 + randUpper],
                [372.5,540],
            ],

            "anchors_pagina_4_2_upper":[
                [res_x,res_y],
                [x_3,y_4 + randUpper],
                [x_3,y_3 + randLower],
                [x_2,y_3 + randUpper],
                [372.5,540],
            ],

            "anchors_pagina_4_2_lower":[
                [res_x,res_y],
                [x_3,y_4 + randLower],
                [x_3,y_3 + randUpper],
                [x_2,y_3 + randLower],
                [372.5,540],
            ],

            "anchors_pagina_4_3_upper":[
                [res_x,res_y],
                [x_4,y_4 + randUpper],
                [x_3,y_4 + randLower],
                [x_3,y_3 + randUpper],
                [x_2,y_3 + randLower],
                [372.5,540],
            ],

            "anchors_pagina_4_3_lower":[
                [res_x,res_y],
                [x_4,y_4 + randLower],
                [x_3,y_4 + randUpper],
                [x_3,y_3 + randLower],
                [x_2,y_3 + randUpper],
                [372.5,540],
            ],

            "anchors_pagina_4_4_upper":[
                [res_x,res_y],
                [x_5,y_4 + randUpper],
                [x_4,y_4 + randLower],
                [x_3,y_4 + randUpper],
                [x_3,y_3 + randLower],
                [x_2,y_3 + randUpper],
                [372.5,540],
            ],

            "anchors_pagina_4_4_lower":[
                [res_x,res_y],
                [x_5,y_4 + randLower],
                [x_4,y_4 + randUpper],
                [x_3,y_4 + randLower],
                [x_3,y_3 + randUpper],
                [x_2,y_3 + randLower],
                [372.5,540],
            ],
        };
    

        //Crea una linea
        var gl = doc.graphicLines.add();
            //Definisce i punti di ancoraggio della linea a seconda che la parola si trovi nella:

        
            //Pagina 1_1
            if(res_x > pw && res_x < pw * 2 && res_y > 0 && res_y < ph /2) {
                for (var j in data.anchors_pagina_1_1_upper) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_1_1_upper[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_1_1_upper[j];            
                    }
                }
            }
            if(res_x > pw && res_x < pw * 2 && res_y > ph / 2 && res_y < ph) {
                for (var j in data.anchors_pagina_1_1_lower) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_1_1_lower[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_1_1_lower[j];            
                    }
                }
            }

            //Pagina 1_2
            if(res_x > pw * 2 && res_x < pw * 3 && res_y > 0 && res_y < ph / 2) {
                for (var j in data.anchors_pagina_1_2_upper) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_1_2_upper[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_1_2_upper[j];            
                    }
                }
            }
            if(res_x > pw * 2 && res_x < pw * 3 && res_y > ph / 2 && res_y < ph) {
                for (var j in data.anchors_pagina_1_2_lower) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_1_2_lower[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_1_2_lower[j];            
                    }
                }
            }

            //Pagina 1_3
            if(res_x > pw * 3 && res_x < pw * 4 && res_y > 0 && res_y < ph / 2) {
                for (var j in data.anchors_pagina_1_3_upper) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_1_3_upper[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_1_3_upper[j];            
                    }
                }
            }
            if(res_x > pw * 3 && res_x < pw * 4 && res_y > ph / 2 && res_y < ph) {
                for (var j in data.anchors_pagina_1_3_lower) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_1_3_lower[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_1_3_lower[j];            
                    }
                }
            }

            //Pagina 2_1
            if(res_x > 0 && res_x < pw && res_y > ph && res_y < ph * 2 + ph / 2) {
                for (var j in data.anchors_pagina_2_1_upper) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_2_1_upper[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_2_1_upper[j];            
                    }
                }
            }
            if(res_x > 0 && res_x < pw && res_y > ph + ph / 2 && res_y < ph * 2) {
                for (var j in data.anchors_pagina_2_1_lower) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_2_1_lower[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_2_1_lower[j];            
                    }
                }
            }

            //Pagina 2_2
            if(res_x > pw && res_x < pw * 2 && res_y > ph && res_y < ph + ph / 2) {
                for (var j in data.anchors_pagina_2_2_upper) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_2_2_upper[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_2_2_upper[j];            
                    }
                }
            }
            if(res_x > pw && res_x < pw * 2 && res_y > ph + ph / 2 && res_y < ph * 2) {
                for (var j in data.anchors_pagina_2_2_lower) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_2_2_lower[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_2_2_lower[j];            
                    }
                }
            }

            //Pagina 2_3
            if(res_x > pw * 2 && res_x < pw * 3 && res_y > ph && res_y < ph + ph / 2) {
                for (var j in data.anchors_pagina_2_3_upper) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_2_3_upper[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_2_3_upper[j];            
                    }
                }
            }
            if(res_x > pw * 2 && res_x < pw * 3 && res_y > ph + ph / 2 && res_y < ph * 2) {
                for (var j in data.anchors_pagina_2_3_lower) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_2_3_lower[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_2_3_lower[j];            
                    }
                }
            }

            //Pagina 2_4
            if(res_x > pw * 3 && res_x < pw * 4 && res_y > ph && res_y < ph + ph / 2) {
                for (var j in data.anchors_pagina_2_4_upper) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_2_4_upper[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_2_4_upper[j];            
                    }
                }
            }
            if(res_x > pw * 3 && res_x < pw * 4 && res_y > ph + ph / 2 && res_y < ph * 2) {
                for (var j in data.anchors_pagina_2_4_lower) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_2_4_lower[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_2_4_lower[j];            
                    }
                }
            }

            //Pagina 3_1
            if(res_x > pw && res_x < pw * 2 && res_y > ph * 2 && res_y < ph * 2 + ph / 2) {
                for (var j in data.anchors_pagina_3_1_upper) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_3_1_upper[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_3_1_upper[j];            
                    }
                }
            }
            if(res_x > pw && res_x < pw * 2 && res_y > ph * 2 + ph / 2 && res_y < ph * 3) {
                for (var j in data.anchors_pagina_3_1_lower) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_3_1_lower[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_3_1_lower[j];            
                    }
                }
            }

            //Pagina 3_2
            if(res_x > pw * 2 && res_x < pw * 3 && res_y > ph * 2 && res_y < ph * 2 + ph / 2) {
                for (var j in data.anchors_pagina_3_2_upper) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_3_2_upper[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_3_2_upper[j];            
                    }
                }
            }
            if(res_x > pw * 2 && res_x < pw * 3 && res_y > ph * 2 + ph / 2 && res_y < ph * 3) {
                for (var j in data.anchors_pagina_3_2_lower) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_3_2_lower[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_3_2_lower[j];            
                    }
                }
            }

            //Pagina 3_3
            if(res_x > pw * 3 && res_x < pw * 4 && res_y > ph * 2 && res_y < ph * 2 + ph / 2) {
                for (var j in data.anchors_pagina_3_3_upper) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_3_3_upper[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_3_3_upper[j];            
                    }
                }
            }
            if(res_x > pw * 3 && res_x < pw * 4 && res_y > ph * 2 + ph / 2 && res_y < ph * 3) {
                for (var j in data.anchors_pagina_3_3_lower) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_3_3_lower[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_3_3_lower[j];            
                    }
                }
            }

            //Pagina 4_1
            if(res_x > pw && res_x < pw * 2 && res_y > ph * 3 && res_y < ph * 3 + ph / 2) {
                for (var j in data.anchors_pagina_4_1_upper) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_4_1_upper[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_4_1_upper[j];            
                    }
                }
            }
            if(res_x > pw && res_x < pw * 2 && res_y > ph * 3 + ph / 2 && res_y < ph * 4) {
                for (var j in data.anchors_pagina_4_1_lower) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_4_1_lower[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_4_1_lower[j];            
                    }
                }
            }

            //Pagina 4_2
            if(res_x > pw * 2 && res_x < pw * 3 && res_y > ph * 3 && res_y < ph * 3 + ph / 2) {
                for (var j in data.anchors_pagina_4_2_upper) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_4_2_upper[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_4_2_upper[j];            
                    }
                }
            }
            if(res_x > pw * 2 && res_x < pw * 3 && res_y > ph * 3 + ph / 2 && res_y < ph * 4) {
                for (var j in data.anchors_pagina_4_2_lower) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_4_2_lower[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_4_2_lower[j];            
                    }
                }
            }

            //Pagina 4_3
            if(res_x > pw * 3 && res_x < pw * 4 && res_y > ph * 3 && res_y < ph * 3 + ph / 2) {
                for (var j in data.anchors_pagina_4_3_upper) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_4_3_upper[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_4_3_upper[j];            
                    }
                }
            }
            if(res_x > pw * 3 && res_x < pw * 4 && res_y > ph * 3 + ph / 2 && res_y < ph * 4) {
                for (var j in data.anchors_pagina_4_3_lower) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_4_3_lower[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_4_3_lower[j];            
                    }
                }
            }

            //Pagina 4_4
            if(res_x > pw * 4 && res_x < pw * 5 && res_y > ph * 3 && res_y < ph * 3 + ph / 2) {
                for (var j in data.anchors_pagina_4_4_upper) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_4_4_upper[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_4_4_upper[j];            
                    }
                }
            }
            if(res_x > pw * 4 && res_x < pw * 5 && res_y > ph * 3 + ph / 2 && res_y < ph * 4) {
                for (var j in data.anchors_pagina_4_4_lower) {
                    var point = gl.paths[0].pathPoints[j];
                    if (j < 2) {
                        point.anchor = data.anchors_pagina_4_4_lower[j];
                    }
                    else {
                        point = gl.paths[0].pathPoints.add();
                        point.anchor = data.anchors_pagina_4_4_lower[j];            
                    }
                }
            }
  };
} //fine della funzione "main"

function genRand (min, max, decimalPlaces) {
    var result = Math.random() * (max - min) + min;
    var power = Math.pow(10, decimalPlaces);
    var result = Math.floor(result * power) / power;
    if (result % 2 == 0){
        result = result + 1
    } else {result = result}
    var value = 17.5 * result;
return value;
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