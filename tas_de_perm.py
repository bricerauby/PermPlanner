class Tas_De_Perm:
    #    class qui contient les perms triées par horaire puis par sensibilité (initialemet vide les perms sont ajoutées par ajoute_perm).
    #    On attibuera les perms horaires par horaires puis en matchant la fiabilité a la sensibilité
    def __init__(self, domaine=None):
        self.perms = []
        self.domaine = domaine

    def ajoute_perm(self, perm):
        if not self.perms:
            self.perms = [perm]
        else:
            i = 0
            j = len(self.perms) - 1
            while self.perms[i].heure != self.perms[j].heure and i < j:
                if perm.heure > self.perms[(i + j) // 2].heure:
                    i = (i + j + 1) // 2
                else:
                    j = (i + j) // 2
            if self.perms[i].heure > perm.heure:
                self.perms = self.perms[:i] + [perm] + self.perms[i:]
            elif self.perms[j].heure < perm.heure:
                self.perms = self.perms[:j + 1] + [perm] + self.perms[j + 1:]
            else:
                while self.perms[i].sensibilite != self.perms[j].sensibilite and i != j:
                    if perm.sensibilite > self.perms[(i + j) // 2].sensibilite:
                        i = (i + j + 1) // 2
                    else:
                        j = (i + j) // 2
                if self.perms[i].sensibilite > perm.sensibilite:
                    self.perms = self.perms[:i] + [perm] + self.perms[i:]
                elif self.perms[j].sensibilite < perm.sensibilite:
                    self.perms = self.perms[:j + 1] + [perm] + self.perms[j + 1:]
                else:  # ici on peut rajouter un dernier tri
                    self.perms = self.perms[:i] + [perm] + self.perms[i:]

    def get_one_hour(self):
        if not self.perms:
            return []
        else:
            i = 0
            while i < len(self.perms) and self.perms[i].heure == self.perms[0].heure:
                i += 1
            res = self.perms[:i]
            self.perms = self.perms[i:]
            return res
