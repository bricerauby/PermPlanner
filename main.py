#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from file_de_staffeur import *
from staffeur import Staffeur
from perm import *
from tas_de_perm import *
import os
import shutil
from weasyprint import HTML
import time
import codecs 
from typing import TextIO
import tqdm
# A remplir
premier_jour = 'Vendredi'

jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
dict_jours = {}
i = jours.index(premier_jour)
for j in range(i, i + 7):
    dict_jours[jours[j % 7]] = j
temps_import = time.time()
# bloc initialisation staffeur
file_staff = file_de_staffeur(
    'main')  # on utlise pas le domaine pour le moment (c'est utile si on veut separer en plusieurs domaine: perm
# secu perm acti etc)
staff_file = openpyxl.load_workbook(filename='data/Liste_Staffeurs.xlsx', read_only=True)
ws = staff_file.active
ws.guess_types = False
for row in ws.rows:
    champs = []
    for cell in row:
        if len(champs) < 6:
            champs.append(cell.value)
    if len(champs)>0 and champs[0] is not None and champs[0] != "Nom":
        staff = Staffeur(*champs)
        if staff.competence is None :
            staff.competence = ""
        staff.competence = staff.competence.replace('/','').split()
        file_staff.ajout_staffeur(staff)
# bloc initialisation heure de perm
tas_perm = Tas_De_Perm('main')
for file_name in os.listdir('data/perms'):
    if file_name[-5:] == '.xlsx':
        perms = import_from_excel(os.path.join('data/perms', file_name), dict_jours)
        for i in perms:
            tas_perm.ajoute_perm(i)
temps_import = time.time() - temps_import
print('fin initialisation')
print('temps d\'initialisation: ' + str(temps_import))
# bloc attribution des perms
temps_calcul = time.time()
heure_debut = tas_perm.perms[0].heure
perms_attribuees = []
print('début attribution ...')
while tas_perm.perms:
    perms_courantes = tas_perm.get_one_hour()
    check = False
    i_max = len(perms_courantes)
    perms_restantes = []
    while perms_courantes:
        # bloc pour forcer la competence requise
        if perms_courantes[0].competence_requise is not None:
            i = 0
            staff_ok = False
            while not staff_ok and i < len(file_staff.staffeurs):
                for comp in file_staff.staffeurs[i].competence:
                    if comp == perms_courantes[0].competence_requise:
                        staff_ok = file_staff.staffeurs[i].staffeur_dispo(perms_courantes[0])
                if not staff_ok:
                    i += 1
            if i < len(file_staff.staffeurs) and file_staff.staffeurs[i].staffeur_dispo(perms_courantes[0]):
                staff = file_staff.get_staffeur(i)
                staff.creneaux.append((perms_courantes[0].nom, perms_courantes[0].heure, perms_courantes[0].duree))
                staff.affectation.append(perms_courantes[0])
                staff.set_heure_dispo(perms_courantes[0])
                perms_attribuees.append(perms_courantes[0])
                perms_attribuees[-1].affectation = staff
                perms_courantes = perms_courantes[1:]
                staff.pen_end += perms_attribuees[-1].penibilite
                staff.soiree_niquee = staff.soiree_niquee or perms_attribuees[-1].nique_soiree
                file_staff.ajout_staffeur(staff)
                i_max += -1
            else:
                raise ValueError("Staffeurs imcompetents pour la perm : " + str(perms_courantes[0].nom))
        else:
            # il  n'y donc pas de competences requises
            if perms_courantes[0].competence_recommandee is not None and (not check):
                staff_ok = False
                i = 0
                while (not staff_ok) and i <= i_max:
                    for comp in file_staff.staffeurs[i].competence:
                        if comp == perms_courantes[0].competence_recommandee and not (
                                file_staff.staffeurs[i].soiree_niquee and perms_courantes[0].nique_soiree):
                            staff_ok = True
                    if not staff_ok:
                        i += 1

                if staff_ok and file_staff.staffeurs[i].staffeur_dispo(perms_courantes[0]):
                    staff = file_staff.get_staffeur(i)
                    staff.creneaux.append(
                        (perms_courantes[0].nom, perms_courantes[0].heure, float(perms_courantes[0].duree)))
                    staff.affectation.append(perms_courantes[0])
                    staff.set_heure_dispo(perms_courantes[0])
                    perms_attribuees.append(perms_courantes[0])
                    perms_attribuees[-1].affectation = staff
                    perms_courantes = perms_courantes[1:]
                    staff.pen_end += perms_attribuees[-1].penibilite
                    staff.soiree_niquee = staff.soiree_niquee or perms_attribuees[-1].nique_soiree
                    file_staff.ajout_staffeur(staff)
                    i_max += -1
                    check = False
                else:
                    check = True

            else:
                perms_restantes.append(perms_courantes[0])
                perms_courantes = perms_courantes[1:]
    j_min = 0
    i_max = len(perms_restantes) - 1

    while perms_restantes:
        staffs = file_staff.staffeurs[j_min:i_max + 1 + j_min]
        staffeurs_affectes = []
        for i in range(len(staffs)):
            for j in range(len(staffs)):
                if staffs[i].fiabilite > staffs[j].fiabilite:
                    temp = staffs[i]
                    staffs[i] = staffs[j]
                    staffs[j] = temp
        for j in range(len(staffs)):
            if staffs[j].staffeur_dispo(perms_restantes[0]):
                staffs[j].creneaux.append(
                    (perms_restantes[0].nom, perms_restantes[0].heure, float(perms_restantes[0].duree)))
                staffs[j].affectation.append(perms_restantes[0])
                staffs[j].set_heure_dispo(perms_restantes[0])
                perms_attribuees.append(perms_restantes[0])
                perms_attribuees[-1].affectation = staffs[j]
                perms_restantes = perms_restantes[1:]
                staffs[j].pen_end += perms_attribuees[-1].penibilite
                staffs[j].soiree_niquee = staffs[j].soiree_niquee or perms_attribuees[-1].nique_soiree
                i_max += -1
                check = False
                staffeurs_affectes.append(staffs[j])
        j_min += len(staffs)
        for j in staffeurs_affectes:
            i = 0
            while i < len(file_staff.staffeurs) and j.nom != file_staff.staffeurs[i].nom:
                i += 1
            temp = file_staff.get_staffeur(i)
            file_staff.ajout_staffeur(temp)
            j_min += -1
print('fin attribution')
temps_calcul = time.time() - temps_calcul
print("temps attribution : " + str(temps_calcul))
perms_pour_staffeurs_wb = openpyxl.Workbook()
sht = perms_pour_staffeurs_wb.active
heure_fin = 0
row = 2
temps_export = time.time()
for j in tqdm.tqdm(file_staff.staffeurs):
    col = 1
    sht.cell(column=1, row=row, value=j.nom + ' ' + j.prenom)
    heure_colonne = heure_debut
    col += 1
    for perm in j.creneaux:
        while heure_colonne != perm[1]:
            col += 1
            if heure_colonne % 10000 == 2300:
                heure_colonne += 10000 - 2300
            else:
                heure_colonne += 100
        if heure_colonne > heure_fin:
            heure_fin = heure_colonne
        for i in range(int(float(perm[2]) + 1)):
            if perm[2] - i > 1:
                sht.cell(column=col, row=row, value=perm[0] + ' (1h)')
                col += 1
                if heure_colonne % 10000 == 2300:
                    heure_colonne += 10000 - 2300
                else:
                    heure_colonne += 100
            elif perm[2] - i == 1:
                sht.cell(column=col, row=row, value=perm[0] + ' (1h)')
            elif 0 < perm[2] - i < 1:
                sht.cell(column=col, row=row, value=perm[0] + ' (' + str((perm[2] % 1) * 60) + 'min)')
                col += 1
                if heure_colonne % 10000 == 2300:
                    heure_colonne += 10000 - 2300
                else:
                    heure_colonne += 100
    row += 1
# on s'occupe de la premiere ligne
row = 1
col = 1
sht.cell(row=row, column=col, value='Nom')
heure_colonne = heure_debut
while heure_colonne <= heure_fin:
    col += 1
    sht.cell(row=row, column=col,
             value=jours[(heure_colonne // 10000) % 7] + ' : ' + str((heure_colonne % 10000) // 100) + 'h')
    if heure_colonne % 10000 == 2300:
        heure_colonne += 10000 - 2300
    else:
        heure_colonne += 100

perms_pour_staffeurs_wb.save('data/perms_pour_staffeurs.xlsx')

print('perm pour staff fini')

# on recupere toutes les perms qui ont le meme nom et on les groupe sur une ligne
perms = []
for p in perms_attribuees:
    check = False
    for j in perms:
        if j[0].nom == p.nom:
            j.append(p)
            check = True
    if not check:
        perms.append([p])
staffeurs_pour_perms_wb = openpyxl.Workbook()
sht = staffeurs_pour_perms_wb.active
row = 2
for p in tqdm.tqdm(perms):
    col = 1
    sht.cell(column=col, row=row, value=p[0].nom)
    col += 1
    heure_colonne = heure_debut

    for creneau in p:
        temp = creneau
        while heure_colonne != creneau.heure:
            col += 1
            if heure_colonne % 10000 == 2300:
                heure_colonne += 10000 - 2300
            else:
                heure_colonne += 100
        for i in range(int(float(creneau.duree) + 1)):
            if creneau.duree - i > 1:
                sht.cell(row=row, column=col,
                         value=creneau.affectation.nom + ' ' + creneau.affectation.prenom + ' (1heure)')
                col += 1
                if heure_colonne % 10000 == 2300:
                    heure_colonne += 10000 - 2300
                else:
                    heure_colonne += 100
            elif creneau.duree - i == 1:
                sht.cell(row=row, column=col,
                         value=creneau.affectation.nom + ' ' + creneau.affectation.prenom + ' (1heure)')
            elif 0 < creneau.duree - i < 1:
                sht.cell(row=row, column=col,
                         value=creneau.affectation.nom + ' ' + creneau.affectation.prenom + ' (' + str(
                             (creneau.duree % 1) * 60) + 'min)')
                col += 1
                if heure_colonne % 10000 == 2300:
                    heure_colonne += 10000 - 2300
                else:
                    heure_colonne += 100
    row += 1
row = 1
col = 1
sht.cell(row=row, column=col, value='Nom')
heure_colonne = heure_debut
while heure_colonne <= heure_fin:
    col += 1
    sht.cell(row=row, column=col,
             value=jours[(heure_colonne // 10000) % 7] + ' : ' + str((heure_colonne % 10000) // 100) + 'h')
    if heure_colonne % 10000 == 2300:
        heure_colonne += 10000 - 2300
    else:
        heure_colonne += 100

staffeurs_pour_perms_wb.save('data/staffeurs_pour_perms.xlsx')
print('staffeurs_pour_perms fini')

# creation des carnets de staff
if not os.path.exists("data/Guide_staffeurs"):
    os.mkdir("data/Guide_staffeurs")
if not os.path.exists("data/Guide_staffeurs/html"):
    os.mkdir("data/Guide_staffeurs/html")
if not os.path.exists("data/Guide_staffeurs/images"):
    os.mkdir("data/Guide_staffeurs/images")

for staffeur in tqdm.tqdm(file_staff.staffeurs):
    with codecs.open('data/Guide_staffeurs/html/' + str(staffeur.nom).replace(' ', '') + '_' + str(staffeur.prenom).replace(' ',
                                                                                                                    '') + ".html",
        "w", encoding='utf8') as fichier_html:
        fichier_html.write("<!DOCTYPE html> "
                           "<html> "
                           "<head>"
                           " <title> ")
        fichier_html.write("Fiche staffeur de " + str(staffeur.prenom) + " " + str(staffeur.nom) + "</title>")
        fichier_html.write('<link rel="stylesheet" type="text/css" href="../style.css"> <link rel="stylesheet" '
                           'type="text/css" href="../impression.css" media="print"> <meta http-equiv="Content-Type" '
                           'content="text/html;charset=utf8" /> '
                           '</head> '
                           '<body>')
        fichier_html.write('<div class="page">'
                           '<div class="container">	<img src="../../logo.png" alt="pagedegarde" style="width:100%; '
                           'z-index: -1;"> <div class="text"> <h1 class="pdg">WEI CENTRALESUPELEC 2018</h1> <h2 '
                           'class="pdg"> Carnet de staff de ' + str(staffeur.prenom) + " " + str(
            staffeur.nom) + '</h2> </div> </div> '
                            '</div>')
    ind = 1
    for perm_to_write in staffeur.affectation:
        with codecs.open('data/Guide_staffeurs/html/' + str(staffeur.nom).replace(' ', '') + '_' + str(staffeur.prenom).replace(' ',
                                                                                                                    '') + ".html",
                  "a", encoding='utf8') as fichier_html:

            if ind < len(staffeur.affectation):
                fichier_html.write('<div class="page"> ')
            fichier_html.write('<h1> ' + perm_to_write.nom + ' </h1> ')
            if perm_to_write.image:
                shutil.copyfile(perm_to_write.image,
                                "data/Guide_staffeurs/images/" + str(perm_to_write.nom.replace(' ', '_').replace('.', '_') + '.png'))
                fichier_html.write('<div class ="plan"> '
                                   '<h2> Lieu de la permanence : </h2>'
                                   '<img src="' + "../images/" + str(
                    perm_to_write.nom.replace(' ', '_').replace('.','_')) + '.png' + '" alt="plan_perm"> </div>')
            fichier_html.write(' <table> <tr> <td>Heure de début</td>  <td>' +
                               str(jours[(perm_to_write.heure // 10000) % 7] + ' : ' + str(
                                   (
                                           perm_to_write.heure % 10000) // 100) + 'h') + '</td>  </tr> <tr> '
                                                                                         '<td>Heure de fin</td> '
                                                                                         '<td>')
            heure_fin_perm = perm_to_write.heure
            if heure_fin_perm % 10000 == 2300:
                heure_fin_perm += 10000 - 2300
            else:
                heure_fin_perm += 100

            fichier_html.write(str(jours[(heure_fin_perm // 10000) % 7] + ' : ' + str((
                                                                                              heure_fin_perm % 10000) // 100)) + 'h </td> </tr> <tr> <td>Responsable de la permanence</td> <td>' + str(
                perm_to_write.respo) + '</td> </tr> <tr> <td>Numéro du responsable</td> <td>' + str(
                perm_to_write.num_respo) + '</td> </tr> <tr> <td>Prérequis</td> <td>' + str(
                perm_to_write.prerequis) + '</td> </tr>')
            for i in range(len(perm_to_write.coperms)):
                if i == 0 and len(perm_to_write.coperms) < 2:
                    fichier_html.write(' <tr> <td>Co-staffeur</td>')
                elif i == 0:
                    fichier_html.write(' <tr> <td>Co-staffeurs</td>')
                else:
                    fichier_html.write(' <tr> <td></td>')
                try:
                    fichier_html.write(
                        '<td>' + perm_to_write.coperms[i].affectation.nom + ' ' + perm_to_write.coperms[
                            i].affectation.prenom + '</td> </tr>') #'(' + perm_to_write.coperms[i].affectation.tel + ')</td> </tr>')
                except TypeError:
                    print('erreur')
                    print(perm_to_write.nom)
                    print(perm_to_write.coperms[i].affectation.nom)
                    print(perm_to_write.coperms[i].affectation.prenom)
                    print(perm_to_write.coperms[i].affectation.tel)
            try :
                fichier_html.write(
                '<tr> <td>Materiel Forunis</td> <td> ' + str(perm_to_write.materiel) + '</td> </tr> </table> <div '
                                                                                  'class="description"> '
                                                                                  '<h2>Description :</h2> <p> ' +
                str(
                    perm_to_write.description) + '</p> </div>')
            except TypeError:
                    print('erreur')
                    print(perm_to_write.nom)
                    print(perm_to_write.description)
            if ind < len(staffeur.affectation):
                fichier_html.write('</div>')
        ind += 1
with open('data/Guide_staffeurs/html/' + str(staffeur.nom).replace(' ', '') + '_' + str(staffeur.prenom).replace(' ',
                                                                                                            '') + ".html",
          "a") as fichier_html:
    fichier_html.write('</body> </html>')
with open('data/Guide_staffeurs/style.css', 'w') as style_file:
    style_file.write('h1 {	text-align: center;	margin: auto;	margin-bottom: 30px;} .plan > h2,img{	margin: '
                     'auto;	width: auto;} .plan >h2 {text-align: center;} .plan >img { display: block; margin-left: '
                     'auto; margin-right: auto;    width: 50%;  } .plan { 	margin-bottom: 30px; } table{    '
                     'border-collapse: collapse;  margin: auto; margin-bottom: 30px; } .container {    position: '
                     'relative; } .text {	position: absolute; 	top : 50%; 	left:0%;	right: 0%;	text-align: '
                     'center;} .pdg {	color: #3701fe;} .container > img {     width: 100%;    height: auto;    '
                     'opacity: 0.3;    z-index: -1; } td{    border: 1px solid black;    padding: 5px ;} .description '
                     '{	width : 65%;	margin:auto;} .page {	margin-bottom: 100px;} '
                     )
with open('data/Guide_staffeurs/impression.css', 'w') as impression_file:
    impression_file.write('.page {page-break-after:always;}')
temps_export = time.time() - temps_export
print('exportation fini')
print('temps d\'exportation :' + str(temps_export))
temps_pdf = time.time()


print("creating pdf...")
if not os.path.exists('data/Guide_staffeurs/pdf'):
    os.mkdir("data/Guide_staffeurs/pdf")
for fichier in tqdm.tqdm(os.listdir("data/Guide_staffeurs/html/")):
    file_name = "data/Guide_staffeurs/pdf/" + fichier[:-5] + '.pdf'
    HTML("data/Guide_staffeurs/html/" + fichier).write_pdf(file_name)
temps_pdf = time.time() - temps_pdf

print("temps_import: " + str(temps_import))
print("temps_calcul: " + str(temps_calcul))
print("temps_export: " + str(temps_export))
print("temps_pdf: " + str(temps_pdf))
