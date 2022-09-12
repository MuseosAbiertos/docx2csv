# docx2csv
docx2csv.py parses .docx files to extract artwork information, then attempts to
pair them up with image files residing in the same directories via filename similarity. 
Extract the following tags VRA Core: 'Work Agent', 'Work Title', 'Work ID', 'Work Type', 'Work Description', 'Work Measurements', 'Work Date'
It finally outputs a CSV associating each image file to the data parsed from the .docx sharing its name.

It can be run using: "python docx2csv.py ROOT_PATH", or, alternatively the path will be asked for at runtime.

---------------------------

docx2csv.py analiza los archivos .docx para extraer la información de las ilustraciones,
y luego intenta emparejarlos con los archivos de imagen que residen en los mismos directorios
a través de la similitud de los nombres de los archivos. 
Extrae los siguientes tags VRA Core: 'Work Agent', 'Work Title', 'Work ID', 'Work Type', 'Work Description', 'Work Measurements', 'Work Date'
Finalmente, genera un CSV que asocia cada archivo de imagen a los datos analizados del .docx que comparte su nombre.

Se puede ejecutar utilizando: "python docx2csv.py ROOT_PATH", o, alternativamente, se pedirá la ruta en tiempo de ejecución.
