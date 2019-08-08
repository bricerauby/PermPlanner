class Staffeur:
    # Une classe staffeur qui contient
    # les attributs nom, prenoms, telephone, creneaux,
    # assos, fiabilite, compétences particulieres, et penibilite enduree

    def __init__(self, nom, prenom, fiabilite, assos, competence, tel=None):
        self.nom = nom
        self.prenom = prenom
        self.fiabilite = fiabilite
        self.assos = assos
        self.heure_dispo = 0
        self.competence = competence
        self.creneaux = []
        self.affectation = []
        self.pen_end = 0
        self.tel = tel
        self.soiree_niquee = False

    def staffeur_dispo(self, perm):
        return perm.heure > self.heure_dispo + 20 and not (self.soiree_niquee and perm.nique_soiree)

    def set_heure_dispo(self, perm):
        if (perm.heure % 1000 + perm.duree) < 2400:
            self.heure_dispo = perm.heure + perm.duree
        else:
            self.heure_dispo = perm.heure + 10000 - 2400 + perm.duree
