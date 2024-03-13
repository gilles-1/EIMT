Sub ComparerEtAttribuerOptimise()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dictC As Object
    Dim valeurColD As Variant
    
    ' Définir la feuille de travail
    Set ws = ThisWorkbook.Sheets("EIMT_DATA_2022_2023_CSV (versio")
    
    ' Trouver la dernière ligne avec des données dans la colonne A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Initialiser un dictionnaire pour stocker les correspondances de la colonne C
    Set dictC = CreateObject("Scripting.Dictionary")
    
    ' Parcourir chaque ligne de la colonne C et stocker les indices dans le dictionnaire
    For i = 1 To lastRow
        If Not dictC.Exists(ws.Cells(i, 10).Value) Then
            dictC.Add ws.Cells(i, 10).Value, i
        End If
    Next i
    
    ' Parcourir chaque ligne de la colonne A
    For i = 1 To lastRow
        ' Vérifier si la valeur de la colonne A existe dans le dictionnaire
        If dictC.Exists(ws.Cells(i, 1).Value) Then
            ' Récupérer l'indice de la colonne C correspondante
            Dim indiceColC As Long
            indiceColC = dictC(ws.Cells(i, 1).Value)
            
            ' Récupérer la valeur de la colonne D
            valeurColD = ws.Cells(indiceColC, 11).Value
            
            ' Attribuer la valeur de la colonne D à la colonne B à côté de la première colonne
            ws.Cells(i, 2).Value = valeurColD
        End If
    Next i
End Sub


//////////////////////////////////////////////////////////////////////


Sub ComparerEtAttribuerOptimiseV()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dictC As Object
    Dim valeurColD As Variant
    
    ' Définir la feuille de travail
    Set ws = ThisWorkbook.Sheets("EIMT_DATA_2022_2023_CSV (versio")
    
    ' Trouver la dernière ligne avec des données dans la colonne A
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Initialiser un dictionnaire pour stocker les correspondances de la colonne C
    Set dictC = CreateObject("Scripting.Dictionary")
    
    ' Parcourir chaque ligne de la colonne C et stocker les indices dans le dictionnaire
    For i = 1 To lastRow
        If Not dictC.Exists(ws.Cells(i, 13).Value) Then
            dictC.Add ws.Cells(i, 13).Value, i
        End If
    Next i
    
    ' Parcourir chaque ligne de la colonne A
    For i = 1 To lastRow
        ' Vérifier si la valeur de la colonne A existe dans le dictionnaire
        If dictC.Exists(ws.Cells(i, 3).Value) Then
            ' Récupérer l'indice de la colonne C correspondante
            Dim indiceColC As Long
            indiceColC = dictC(ws.Cells(i, 3).Value)
            
            ' Récupérer la valeur de la colonne D
            valeurColD = ws.Cells(indiceColC, 14).Value
            
            ' Attribuer la valeur de la colonne D à la colonne B à côté de la première colonne
            ws.Cells(i, 4).Value = valeurColD
        End If
    Next i
End Sub


///////////////////////////////////////////////////////////////////////////////////////


Sub ComparerEtAttribuerOptimiseEMP()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dictC As Object
    Dim valeurColD As Variant
    
    ' Définir la feuille de travail
    Set ws = ThisWorkbook.Sheets("EIMT_DATA_2022_2023_CSV (versio")
    
    ' Trouver la dernière ligne avec des données dans la colonne A
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    ' Initialiser un dictionnaire pour stocker les correspondances de la colonne C
    Set dictC = CreateObject("Scripting.Dictionary")
    
    ' Parcourir chaque ligne de la colonne C et stocker les indices dans le dictionnaire
    For i = 1 To lastRow
        If Not dictC.Exists(ws.Cells(i, 18).Value) Then
            dictC.Add ws.Cells(i, 18).Value, i
        End If
    Next i
    
    ' Parcourir chaque ligne de la colonne A
    For i = 1 To lastRow
        ' Vérifier si la valeur de la colonne A existe dans le dictionnaire
        If dictC.Exists(ws.Cells(i, 5).Value) Then
            ' Récupérer l'indice de la colonne C correspondante
            Dim indiceColC As Long
            indiceColC = dictC(ws.Cells(i, 5).Value)
            
            ' Récupérer la valeur de la colonne D
            valeurColD = ws.Cells(indiceColC, 19).Value
            
            ' Attribuer la valeur de la colonne D à la colonne B à côté de la première colonne
            ws.Cells(i, 6).Value = valeurColD
        End If
    Next i
End Sub


////////////////////////////////////////////////////////////////////////////////////////////




Sub ComparerEtAttribuerOptimisePROF()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dictC As Object
    Dim valeurColD As Variant
    
    ' Définir la feuille de travail
    Set ws = ThisWorkbook.Sheets("EIMT_DATA_2022_2023_CSV (versio")
    
    ' Trouver la dernière ligne avec des données dans la colonne A
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    ' Initialiser un dictionnaire pour stocker les correspondances de la colonne C
    Set dictC = CreateObject("Scripting.Dictionary")
    
    ' Parcourir chaque ligne de la colonne C et stocker les indices dans le dictionnaire
    For i = 1 To lastRow
        If Not dictC.Exists(ws.Cells(i, 17).Value) Then
            dictC.Add ws.Cells(i, 17).Value, i
        End If
    Next i
    
    ' Parcourir chaque ligne de la colonne A
    For i = 1 To lastRow
        ' Vérifier si la valeur de la colonne A existe dans le dictionnaire
        If dictC.Exists(ws.Cells(i, 8).Value) Then
            ' Récupérer l'indice de la colonne C correspondante
            Dim indiceColC As Long
            indiceColC = dictC(ws.Cells(i, 8).Value)
            
            ' Récupérer la valeur de la colonne D
            valeurColD = ws.Cells(indiceColC, 18).Value
            
            ' Attribuer la valeur de la colonne D à la colonne B à côté de la première colonne
            ws.Cells(i, 9).Value = valeurColD
        End If
    Next i
End Sub




Sub ComparerEtAttribuerOptimiseADR()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dictC As Object
    Dim valeurColD As Variant
    
    ' Définir la feuille de travail
    Set ws = ThisWorkbook.Sheets("EIMT_DATA_2022_2023_CSV (versio")
    
    ' Trouver la dernière ligne avec des données dans la colonne A
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ' Initialiser un dictionnaire pour stocker les correspondances de la colonne C
    Set dictC = CreateObject("Scripting.Dictionary")
    
    ' Parcourir chaque ligne de la colonne C et stocker les indices dans le dictionnaire
    For i = 1 To lastRow
        If Not dictC.Exists(ws.Cells(i, 23).Value) Then
            dictC.Add ws.Cells(i, 23).Value, i
        End If
    Next i
    
    ' Parcourir chaque ligne de la colonne A
    For i = 1 To lastRow
        ' Vérifier si la valeur de la colonne A existe dans le dictionnaire
        If dictC.Exists(ws.Cells(i, 7).Value) Then
            ' Récupérer l'indice de la colonne C correspondante
            Dim indiceColC As Long
            indiceColC = dictC(ws.Cells(i, 7).Value)
            
            ' Récupérer la valeur de la colonne D
            valeurColD = ws.Cells(indiceColC, 24).Value
            
            ' Attribuer la valeur de la colonne D à la colonne B à côté de la première colonne
            ws.Cells(i, 8).Value = valeurColD
        End If
    Next i
End Sub


///////////////////////////////////////////////////////////

Sub ComparerADR1()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dictC As Object
    Dim valeurColD As Variant
    
    ' Définir la feuille de travail
    Set ws = ThisWorkbook.Sheets("ADRESSE1_1_10000 (1)")
    
    ' Trouver la dernière ligne avec des données dans la colonne B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Initialiser un dictionnaire pour stocker les correspondances de la colonne C
    Set dictC = CreateObject("Scripting.Dictionary")
    
    ' Parcourir chaque ligne de la colonne F et stocker les indices dans le dictionnaire
    For i = 1 To lastRow
        If Not dictC.Exists(ws.Cells(i, 6).Value) Then
            dictC.Add ws.Cells(i, 6).Value, i
        End If
    Next i
    
    ' Parcourir chaque ligne de la colonne B
    For i = 1 To lastRow
        ' Vérifier si la valeur de la colonne B existe dans le dictionnaire
        If dictC.Exists(ws.Cells(i, 2).Value) Then
            ' Récupérer l'indice de la colonne B correspondante
            Dim indiceColC As Long
            indiceColC = dictC(ws.Cells(i, 2).Value)
            
            ' Récupérer la valeur de la colonne D
            valeurColD = ws.Cells(indiceColC, 7).Value
            
            ' Attribuer la valeur de la colonne D à la colonne B à côté de la première colonne
            ws.Cells(i, 3).Value = valeurColD
        End If
    Next i
End Sub


/////////////////////////////////////////////////////////////////////////

Sub SeparerCoordonnees()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim coord As Variant
    Dim lat As Double
    Dim lon As Double
    
    ' Spécifiez le nom de votre feuille de calcul
    Set ws = ThisWorkbook.Sheets("ADRESSE")
    
    ' Trouver la dernière ligne avec des données dans la colonne Latlong
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Boucle à travers chaque ligne avec des données
    For i = 1 To lastRow
        ' Séparer les coordonnées
        coord = Split(ws.Cells(i, "C").Value, ",")
        
        ' Vérifier si les coordonnées sont valides
        If UBound(coord) = 1 Then
            ' Convertir les coordonnées en nombres
            lat = CDbl(Trim(coord(0)))
            lon = CDbl(Trim(coord(1)))
            
            ' Mettre à jour les colonnes Latitude et Longitude
            ws.Cells(i, "D").Value = lat
            ws.Cells(i, "E").Value = lon
        End If
    Next i
End Sub

TCHGI0609001

https://trac.osgeo.org/postgis/wiki/UsersWikiPostGIS3UbuntuPGSQLApt
https://www.paulshapley.com/2022/12/install-postresql-14-and-postgis-3-on.html
https://tecadmin.net/install-postgis-on-ubuntu/

https://docs.geoserver.org/main/en/user/installation/linux.html
https://www.linkedin.com/pulse/install-geoserver-ubuntu-server-krishna-lodha/
https://www.linkedin.com/pulse/install-geoserver-ubuntu-server-krishna-lodha/

https://www.youtube.com/watch?v=D13ZeWvFoT4

https://www.digitalocean.com/community/tutorials/how-to-install-apache-tomcat-10-on-ubuntu-20-04

https://www.hostinger.fr/tutoriels/comment-installer-tomcat-sur-ubuntu#:~:text=La%20meilleure%20fa%C3%A7on%20d'installer,suivez%20la%20derni%C3%A8re%20version%20stable.

https://www.hostinger.fr/tutoriels/comment-installer-tomcat-sur-ubuntu#:~:text=La%20meilleure%20fa%C3%A7on%20d'installer,suivez%20la%20derni%C3%A8re%20version%20stable.

https://medium.com/@DevOpsfreak/how-to-change-the-default-port-of-apache-tomcat-in-ubuntu-and-red-hat-os-a-step-by-step-guide-fb01edaae260
https://www.it-connect.fr/changer-le-port-decoute-de-tomcat/

pg_dump -U [utilisateur] -h [hôte] -p [port] -d [nom_de_la_base] -F c -f [chemin/vers/fichier_de_sortie.dump]

pg_dump -U monutilisateur -h localhost -p 5432 -d mabase -F c -f /chemin/vers/backup.dump

C:\Program Files\Odoo 17.0.20240129\PostgreSQL\bin\pg_dump.exe --file "C:\\Users\\Admin\\OneDrive\\DOCUME~1\\EIMT" --host "localhost" --port "5432" --username "gilles" --no-password --verbose --role "gilles" --format=c --blobs "EIMT"


pg_restore -U [utilisateur] -h [hôte] -p [port] -d [nom_de_la_base] -F c -c [chemin/vers/fichier_de_sauvegarde.dump]


CREATE EXTENSION IF NOT EXISTS postgis;
ALTER TABLE adresse
ADD COLUMN geom geography(Point, 4326);


-- Mettre à jour la colonne geom avec les valeurs de latitude et longitude
UPDATE adresse
SET geom = ST_SetSRID(ST_MakePoint(longitude, latitude), 4326);

////////////////////////////////////////////////////////////////////////////////////////////////

-- Création de la vue
CREATE OR REPLACE VIEW ma_vue AS
SELECT
    t1.colonne1 AS t1_colonne1,
    t1.colonne2 AS t1_colonne2,
    t2.colonne1 AS t2_colonne1,
    t2.colonne2 AS t2_colonne2,
    t3.colonne1 AS t3_colonne1,
    t3.colonne2 AS t3_colonne2
FROM
    table1 t1
    INNER JOIN table2 t2 ON t1.id = t2.id
    LEFT JOIN table3 t3 ON t1.id = t3.id;
	
	
psql -h host -d database -U user -c "SELECT ST_AsGeoJSON(geom)::json AS geometry, other_column1, other_column2 FROM ma_table" -o output.geojson

https://www.youtube.com/watch?v=R-V7XFUbrkw

https://www.youtube.com/watch?v=WnPcSGlh0eQ                 LEAFLET CONTROL SEARCH

test d'habiletés cognitives Emploi de niveau professionnel

test d'habiletés cognitives Emploi d'analyste de l'informatique

https://www.youtube.com/watch?v=FWgrhnTu7Yo&list=PLPBe3S1JOod1STE1ROrJ8KTkCtdTcSst6
https://www.youtube.com/watch?v=nWf4qDva2kk

https://www.youtube.com/watch?v=QGuEKT2Pdb8           QUEBEC   BON


psql -h localhost -d EIMT -U gilles -c "SELECT annee, province_territoire, volet_progranne, profession, employeur, postes_approuves, eimt_approuves, adresse, ST_AsGeoJSON(geom)::json AS geom FROM fait_approbations INNER JOIN adresse ON fait_approbations.adresse_id = adresse.adresse_id, INNER JOIN annee ON fait_approbations.annee_id = annee.annee_id, INNER JOIN employeur ON fait_approbations.employeur_id = employeur.employeur_id, INNER JOIN profession ON fait_approbations.profession_id = profession.profession_id, INNER JOIN province_territoire ON fait_approbations.province_territoire_id = province_territoire.province_territoire_id, INNER JOIN volet_programme ON fait_approbations.volet_programme_id = volet_programme.volet_programme_id" -o data.geojson


psql -h localhost -d EIMT -U gilles -c "SELECT annee, province_territoire, volet_programme, profession, employeur, postes_approuves, eimt_approuves, adresse, ST_AsGeoJSON(geom)::json AS geom FROM fait_approbations INNER JOIN adresse ON fait_approbations.adresse_id = adresse.adresse_id INNER JOIN annee ON fait_approbations.annee_id = annee.annee_id INNER JOIN employeur ON fait_approbations.employeur_id = employeur.employeur_id INNER JOIN profession ON fait_approbations.profession_id = profession.profession_id INNER JOIN province_territoire ON fait_approbations.province_territoire_id = province_territoire.province_territoire_id INNER JOIN volet_programme ON fait_approbations.volet_programme_id = volet_programme.volet_programme_id" -o data1.geojson

ogr2ogr -f "GeoJSON" data1.geojson data.geojson

# Exécutez la requête SQL et enregistrez la sortie dans un fichier GeoJSON
psql -h localhost -d EIMT -U gilles -t -c "SELECT row_to_json(fc) FROM (SELECT 'FeatureCollection' As type, array_to_json(array_agg(f)) As features FROM (SELECT 'Feature' As type, ST_AsGeoJSON(geom)::json As geometry, row_to_json((annee, province_territoire, volet_programme, profession, employeur, postes_approuves, eimt_approuves, adresse)) As properties FROM fait_approbations INNER JOIN adresse ON fait_approbations.adresse_id = adresse.adresse_id INNER JOIN annee ON fait_approbations.annee_id = annee.annee_id INNER JOIN employeur ON fait_approbations.employeur_id = employeur.employeur_id INNER JOIN profession ON fait_approbations.profession_id = profession.profession_id INNER JOIN province_territoire ON fait_approbations.province_territoire_id = province_territoire.province_territoire_id INNER JOIN volet_programme ON fait_approbations.volet_programme_id = volet_programme.volet_programme_id) As f) As fc" -o data.geojson

# Convertissez le fichier GeoJSON en utilisant ogr2ogr (assurez-vous que GDAL/OGR est installé)
ogr2ogr -f "GeoJSON" output.geojson data.geojson


psql -h localhost -d EIMT -U gilles -t -c "SELECT ST_AsGeoJSON(geom)::json, row_to_json(annee), row_to_json(province_territoire), row_to_json(volet_programme), row_to_json(profession), row_to_json(employeur), row_to_json(postes_approuves), row_to_json(eimt_approuves), row_to_json(adresse) FROM fait_approbations INNER JOIN adresse ON fait_approbations.adresse_id = adresse.adresse_id INNER JOIN annee ON fait_approbations.annee_id = annee.annee_id INNER JOIN employeur ON fait_approbations.employeur_id = employeur.employeur_id INNER JOIN profession ON fait_approbations.profession_id = profession.profession_id INNER JOIN province_territoire ON fait_approbations.province_territoire_id = province_territoire.province_territoire_id INNER JOIN volet_programme ON fait_approbations.volet_programme_id = volet_programme.volet_programme_id" -o data.geojson

psql -h localhost -d EIMT -U gilles -t -c "SELECT array_to_json(array_agg(f)) As features FROM (SELECT 'Feature' As type, ST_AsGeoJSON(geom)::json As geometry, row_to_json((annee, province_territoire, volet_programme, profession, employeur, postes_approuves, eimt_approuves, adresse)) As properties FROM fait_approbations INNER JOIN adresse ON fait_approbations.adresse_id = adresse.adresse_id INNER JOIN annee ON fait_approbations.annee_id = annee.annee_id INNER JOIN employeur ON fait_approbations.employeur_id = employeur.employeur_id INNER JOIN profession ON fait_approbations.profession_id = profession.profession_id INNER JOIN province_territoire ON fait_approbations.province_territoire_id = province_territoire.province_territoire_id INNER JOIN volet_programme ON fait_approbations.volet_programme_id = volet_programme.volet_programme_id) As f" -o data.geojson


