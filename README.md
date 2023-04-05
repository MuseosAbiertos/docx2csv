# docx2csv
The "docx2csv" script is a Python tool that allows for the conversion of Microsoft Word documents (.docx) to CSV (Comma-Separated Values) files, which makes it easier to process them in other applications and tools.

The script uses the python-docx library to read the content of the .docx file and the csv library to write the data to the resulting CSV file. It also includes options to select the field delimiter and quote character.

Using the script is very simple; one only needs to run the file with the input .docx file name and the output .csv file name as arguments on the command line. The output file will be created in the same directory as the input file.

This script can be useful for those who need to process large amounts of data in Word document format and want to convert them to a more manageable format.
Extract the following tags VRA Core: 'Work Agent', 'Work Title', 'Work ID', 'Work Type', 'Work Description', 'Work Measurements', 'Work Date'
It finally outputs a CSV associating each image file to the data parsed from the .docx sharing its name.

It can be run using: "python docx2csv.py ROOT_PATH", or, alternatively the path will be asked for at runtime.

---------------------------

El script "docx2csv" es una herramienta en Python que permite convertir documentos de Microsoft Word (.docx) en archivos CSV (valores separados por comas), lo que facilita su procesamiento en otras aplicaciones y herramientas.

El script utiliza la biblioteca python-docx para leer el contenido del archivo .docx y la biblioteca csv para escribir los datos en el archivo CSV resultante. También incluye opciones para seleccionar el delimitador de campo y el carácter de cita.

El uso del script es muy sencillo, simplemente se debe ejecutar el archivo con el nombre del archivo de entrada .docx y el nombre del archivo de salida .csv como argumentos en la línea de comandos. El archivo de salida se creará en el mismo directorio que el archivo de entrada.

Este script puede ser útil para aquellos que necesitan procesar grandes cantidades de datos en formato de documento de Word y desean convertirlos en un formato más manejable.

Extrae los siguientes tags VRA Core: 'Work Agent', 'Work Title', 'Work ID', 'Work Type', 'Work Description', 'Work Measurements', 'Work Date'
Finalmente, genera un CSV que asocia cada archivo de imagen a los datos analizados del .docx que comparte su nombre.

Se puede ejecutar utilizando: "python docx2csv.py ROOT_PATH", o, alternativamente, se pedirá la ruta en tiempo de ejecución.
