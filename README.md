# ExpPatrimonies
Programme Java permettant d'extraire les patrimoines d'une base de données locale dans un fichier Excel

## Utilisation:
```
java ExpPatrimonies [-dbserver db] [-p chemin vers fichier] [-o fichier] [-d] [-t] 
```
où :
* ```-dbserver prod|pre-prod``` est la référence à la base de données, par défaut désigne la base de données de pre-production. Voir fichier *ExpPatrimonies.prop* (optionnel).
* ```-p chemin vers fichier``` est le chemin vers le fichier Excel. Amorcé à vide par défaut (paramètre optionnel).
* ```-o fichier``` est le nom du fichier Excel qui recevra les sociétés. Amorcé à *patrimonies.xlsx* par défaut (paramètre optionnel).
* ```-d``` le programme s'exécute en mode débug, il est beaucoup plus verbeux. Désactivé par défaut (paramètre optionnel).
* ```-t``` le programme s'exécute en mode test, les transactions en base de données ne sont pas faites. Désactivé par défaut (paramètre optionnel).

## Pré-requis :
- Java 6 ou supérieur.
- JDBC Informix
- JDBC MySql
- Driver MongoDB
- [xmlbeans-2.6.0.jar](https://xmlbeans.apache.org/)
- [commons-collections4-4.1.jar](https://commons.apache.org/proper/commons-collections/download_collections.cgi)

## Fichier des paramètres : 

Ce fichier permet de spécifier les paramètres d'accès aux différentes bases de données.

A adapter selon les implémentations locales.

Ce fichier est nommé : *ExpPatrimonies.prop*.

Le fichier *ExpPatrimonies_Example.prop* est fourni à titre d'exemple.

## Références:

- [API Java Exel POI](http://poi.apache.org/download.html)
- [Tuto Java POI Excel](http://thierry-leriche-dessirier.developpez.com/tutoriels/java/charger-modifier-donnees-excel-2010-5-minutes/)
- [Tuto Java POI Excel](http://jmdoudoux.developpez.com/cours/developpons/java/chap-generation-documents.php)
