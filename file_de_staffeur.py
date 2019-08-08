

class file_de_staffeur:
    #    une classe qui contient les staffeurs triés
    #    par sensibilité endurée, actualisée à chaque attribution de perms
    def __init__(self, domaine=None):
        self.staffeurs = []
        self.domaine = domaine

    def ajout_staffeur(self, staffeur):
        if not self.staffeurs:
            self.staffeurs = [staffeur]
        else:
            i = 0
            j = len(self.staffeurs) - 1
            while self.staffeurs[i].pen_end != self.staffeurs[j].pen_end:
                if staffeur.pen_end > self.staffeurs[(i + j) // 2].pen_end:
                    i = (i + j + 1) // 2

                else:
                    j = (i + j) // 2
            if self.staffeurs[i].pen_end > staffeur.pen_end:
                self.staffeurs = self.staffeurs[:i] + [staffeur] + self.staffeurs[i:]
            else:
                self.staffeurs = self.staffeurs[:j + 1] + [staffeur] + self.staffeurs[j + 1:]

    def get_staffeur(self, i):
        if self.staffeurs:
            res = self.staffeurs[i]
            self.staffeurs = self.staffeurs[:i] + self.staffeurs[i + 1:]
            return res
