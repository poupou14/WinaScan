# En cas général les "#" servent à faire des commentaires comme ici
echo Lancement scrap de WinaScan
date > ./date1.txt
python ./src/WS.py $@
date > ./date2.txt
 
