<h1>Instruccions</h1>

1. Crear i configurar les carpetes a l'arxiu ```config.json```:

    ```carpetaSignatures```:  Imatge de les signatures in format jpeg, png, jpg

    ```carpetaDocumentsPeticions```: Document d'al.legacions

    ```csvPeticions```: Llistat de les peticions a fer per persona

    ```csvSignatures```: Llistat de dades personals de les persones signants

    ```carpetaPDFs```: Carpeta on hi ha els PDFs amb els documents d'al.legacions

    ```carpetaDOCXs```: Carpeta on hi ha el DOCX de les al.legacions

    ```carpetaPDFRebuts```: Carpeta on guardar els rebuts després de completar la petició
    

2. Confirmar que tots els placeholders al document word siguin correctes
3. L'script ```CreaPDFs.py``` agafa el word amb els placeholders i en crea un PDF completant el word original amb les dades dels signants. Aquest serà el document d'al.legacions.
4. L'script ```PeticioGeneralitat.py``` accedeix a la web de la Generalitat i per cada al.legació en crea una petició en nom del signant i es descarrega el rebut