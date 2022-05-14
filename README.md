# docx2csv
docx2csv.py parses .docx files to extract artwork information, then attempts to pair them up with image files residing in the same directories via filename similarity. It finally outputs a CSV associating each image file to the data parsed from the .docx sharing its name.

It can be run using: "python docx2csv.py ROOT_PATH", or, alternatively the path will be asked for at runtime.

The script expects the directory structure to only have 1 level of depth, meaning it should be:

root_path
    - directory 1
        - file 1
        - file 2
    - directory 2
        - file 1
        - file 2
--
docx2csv.py analiza los archivos .docx para extraer la información de las ilustraciones, y luego intenta emparejarlos con los archivos de imagen
que residen en los mismos directorios a través de la similitud de los nombres de los archivos. Finalmente, genera un CSV que asocia cada archivo de imagen a los datos analizados del .docx que comparte su nombre.

Se puede ejecutar utilizando: "python docx2csv.py ROOT_PATH", o, alternativamente, se pedirá la ruta en tiempo de ejecución.

El script espera que la estructura de directorios sólo tenga un nivel de profundidad, es decir, que sea:

root_path
    - directory 1
        - file 1
        - file 2
    - directory 2
        - file 1
        - file 2

