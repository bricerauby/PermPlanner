import openpyxl
import os


class Perm:

    # Une classe perm qui contient
    # les attributs nom, heure, duree,
    # sensibilite, penibilite, competences
    # requises, competence recommandee, Lieu,affectation
    # et le fait de niquer la soiree du staffeur
    def __init__(self, nom=None, heure=None, duree=1, sensibilite=5, penibilite=5, competence_requise=None,
                 competence_recommandee=None, image=None,
                 nique_soiree=False, respo=None, num_respo=None, materiel=None, description=None, prerequis=None):
        self.nom = nom
        self.heure = heure
        self.duree = duree
        self.sensibilite = sensibilite
        self.penibilite = penibilite
        self.competence_requise = competence_requise
        self.competence_recommandee = competence_recommandee
        self.affectation = None
        self.nique_soiree = nique_soiree
        assert isinstance(image, object)
        self.image = image
        self.respo = respo
        self.num_respo = num_respo
        self.materiel = materiel
        self.description = description
        self.prerequis = prerequis
        self.coperms = []


def import_from_excel(filename, dict_jours):
    # function that return a list of perm for an excel file
    image_path = filename[:-5] + '.png'
    perm_file = openpyxl.load_workbook(filename=filename, read_only=True)
    ws = perm_file.active
    im_path = None
    print(image_path)
    if os.path.exists(image_path):
        im_path = image_path
    name = filename[:-5].split('/')[-1]
    competence_requise = None
    for row in ws.iter_rows():
        if row[0].value == 'Nombre de staffeurs ':
            nombre_de_staffeurs = int(row[1].value)
        if row[0].value == 'Heure de début':
            jour_debut = dict_jours[row[1].value.replace(' ','').capitalize()]
            heure_debut = int(row[2].value.replace(' ','').replace('h',''))
        if row[0].value == 'Heure de fin':
            jour_fin = dict_jours[row[1].value.replace(' ','').capitalize()]
            heure_fin = int(row[2].value.replace(' ','').replace('h',''))
        if row[0].value == 'Responsable de la permanence':
            respo = row[1].value
        if row[0].value == 'Numéro du responsable':
            num_respo = row[1].value
        if row[0].value == 'Prérequis':
            prerequis = row[1].value
        if row[0].value == 'Fiabilité nécessaire du staffeur':
            sensibilite = int(row[1].value)
        if row[0].value == 'Compétence requise':
            if row[1].value is not None :
                competence_requise = row[1].value.replace('/','')
                if competence_requise == '': 
                    competence_requise = None
            else: 
                competence_requise = row[1].value
        if row[0].value == 'Pénibilité':
            penibilite = int(row[1].value)
        if row[0].value == 'Matériel fournis':
            materiel = row[1].value
        if row[0].value == 'Description':
            description = row[1].value
    perms = []
    jour = jour_debut
    heure = heure_debut
    print(competence_requise)
    while (jour == jour_fin and heure < heure_fin) or jour < jour_fin:
        coperms = []
        for i in range(nombre_de_staffeurs):
            name_perm = name.replace('_', ' ') + '.' + str(i + 1)
            if heure < 10:
                horaire = '0' + str(heure) + '00'
            else:
                horaire = str(heure) + '00'
            if im_path:
                perm = Perm(name_perm, int(str(jour) + str(horaire)), 1, sensibilite=sensibilite, penibilite=penibilite,
                            image=im_path, description=description, prerequis=prerequis, respo=respo,
                            num_respo=num_respo,
                            materiel=materiel,
                            competence_requise=competence_requise)
            else:
                perm = Perm(name_perm, int(str(jour) + str(horaire)), 1, sensibilite=sensibilite, penibilite=penibilite,
                            description=description, prerequis=prerequis, respo=respo, num_respo=num_respo,
                            materiel=materiel,
                            competence_requise=competence_requise)

            coperms.append(perm)
        for perm_ in coperms:
            perm_.coperms = [p for p in coperms if p != perm_]
        perms += coperms
        if heure == 23:
            jour += 1
            heure = 0
        else:
            heure += 1
    return perms
