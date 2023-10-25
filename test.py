import pandas as pd
from typing import Any, Self


def definition(variable: Any):
    taille = 19
    [print(f"{(attrib+' '*taille)[0:taille]} ", end="") for attrib in filter(lambda x: not x.startswith("_"), 
        dir(variable))]
    print()
    # help(pd.DataFrame.drop)
    # exit()


# definition("")
# definition(pd.DataFrame)
# help(pd)
# definition(pd)


class Onglet:

    def __init__(self, sheet_name: str):
        self.rename(sheet_name)

        self.__data: pd.DataFrame = pd.DataFrame()
        self.__nom_cols: list[str] = []
        self.__defaults: list[Any] = []

        self.__nb_cols = 0
        self.__last_index: int = 0

    def rename(self, sheet_name: str) -> None:
        self.__sheet_name = sheet_name        

    def add_colonnes(self, nom_colonnes: list[str], valeurs_defaut: list[Any]) -> None:
        lng_lst1 = len(nom_colonnes)
        lng_lst2 = len(valeurs_defaut)
        if lng_lst1 != lng_lst2:
            raise ValueError(f"Les 2 listes de parametres 'nom_colonnes' et 'valeurs_defaut' doivent avoir le meme nombre de valeurs ({lng_lst1} != {lng_lst2})")

        self.__defaults = valeurs_defaut
        for nom_colonne, valeur_defaut in zip(nom_colonnes, valeurs_defaut):
            self.add_colonne(nom_colonne, valeur_defaut)

    def add_colonne(self, nom_colonne: str, valeur_defaut: Any) -> None:
        if nom_colonne in self.__nom_cols:
            raise ValueError(f"La colonne '{nom_colonne}' existe deja.")

        self.__data[nom_colonne] = valeur_defaut
        self.__nom_cols.append(nom_colonne)
        self.__nb_cols += 1

    def rename_colonne(self, old_name: str, new_name: str) -> None:
        self.__data.rename(columns={old_name: new_name}, inplace=True)
        self.__nom_cols[self.__nom_cols.index(old_name)] = new_name

    def del_colonne(self, colonne_name: str) -> None:
        # self.__data = self.__data.drop(columns=[colonne_name])
        self.__data.drop(columns=[colonne_name], inplace=True)
        self.__nom_cols.remove(colonne_name)
        self.__nb_cols -= 1

    def sort(self, colonnes: list[str]) -> Self:
        self.__data.sort_values(colonnes, inplace=True)
        return self

    def sort_index(self) -> None:
        self.__data.sort_index(inplace=True)

    def reindex(self) -> None:
        self.__data = pd.DataFrame(self.__data.values, 
            columns=self.__data.columns, 
            index=range(1, 1+self.__data.index.shape[0]))
        self.__last_index = self.__data.index.shape[0]
        
    def add_line(self, valeurs: list) -> None:
        nb_val = len(valeurs)
        if nb_val != self.__nb_cols:
            raise ValueError(f"Le nombre de parametres ne correspond pas a la feuille '{nb_val}' au lieu de '{self.__nb_cols}'.")

        self.__last_index += 1
        # print(f"Add: ({self.__last_index}) {valeurs[0]}")

        df: pd.DataFrame = pd.DataFrame([valeurs], columns=self.__nom_cols, index=[self.__last_index])
        self.__data = pd.concat([self.__data, df])

    def del_line(self, line_idx: int) -> None:
        self.__data.drop(index=[line_idx], inplace=True)
        self.__last_index = self.__data.index.max()

    def update_line(self, index: int, **colonnes) -> None:
        for idx, colonne in enumerate(colonnes):
            # print(f"{colonne} set to: {colonnes[colonne]}")
            # self.__data[colonne].iloc[index-1] = colonnes[colonne]
            # self.__data.iloc[index-1, idx] = colonnes[colonne]

            col_index = self.__nom_cols.index(colonne)
            self.__data.iloc[index-1, col_index] = colonnes[colonne]

    def drop_duplicates(self) -> None:
        self.__data.drop_duplicates(inplace=True)

    def getColonnes(self) -> list[str]:
        return self.__nom_cols

    def getData(self) -> pd.DataFrame:
        return self.__data

    def getName(self) -> str:
        return self.__sheet_name

    def to_excel(self, nom_fichier: str) -> None:
        self.__data.to_excel(nom_fichier)

    def __add__(self, other): 
        if "|".join(self.__nom_cols) != "|".join(other.__nom_cols):
            raise ValueError("Le nom des colonnes sont differents")

        nombre1 = 1+self.__data.index.shape[0]
        data1 = pd.DataFrame(self.__data.values, 
            columns=self.__data.columns, 
            index=range(1, nombre1))

        nombre2 = other.__data.index.shape[0]
        data2 = pd.DataFrame(other.__data.values, 
            columns=other.__data.columns, 
            index=range(nombre1, nombre1+nombre2))

        sheet: Onglet = Onglet(self.__sheet_name)
        sheet.add_colonnes(self.__nom_cols, self.__defaults)

        sheet.__data = pd.concat([data1, data2])
        sheet.__last_index = sheet.__data.index.shape[0]

        return sheet
        # return pd.concat([self.__data, other.__data])

    def __str__(self) -> str:
        return f"Onglet: {self.__sheet_name}\n{self.__data}"


class Classeur:

    def __init__(self, filename: str):
        self.rename(filename)
        self.__onglets: list[Onglet] = []

    def rename(self, filename: str) -> None:
        self.__filename = filename

    def add_sheet(self, sheet: Onglet) -> None:
        self.__onglets.append(sheet)

    def drop_sheet(self, sheet: Onglet) -> None:
        self.__onglets.remove(sheet)

    def dropSheetByIndex(self, index: int) -> None:
        del self.__onglets[index]

    def dropSheetByName(self, sheet_name: str) -> None:
        self.__onglets.remove(self.getSheetByName(sheet_name))

    def getSheetsCount(self) -> int:
        return len(self.__onglets)

    def getSheet(self, index=None) -> Onglet:
        if isinstance(index, int):
            return self.getSheetByIndex(index)
        elif isinstance(index, str):
            return self.getSheetByName(index)

        raise Exception(f"Le parametre index est de type '{index.__class__.__name__}' au lieu de int | str.")

    def getSheets(self, param: Any = None) -> Onglet | list[Onglet]:
        if param is None:
            return self.__onglets

        return self.getSheet(param)

    def getSheetByIndex(self, index: int) -> Onglet:
        return self.__onglets[index]

    def getSheetByName(self, sheet_name: str) -> Onglet:
        for sheet in self.__onglets:
            if sheet.getName() == sheet_name:
                return sheet

        raise Exception(f"La feuille '{sheet_name}' n'existe pas dans le Classeur")
        
    def to_excel(self, header: bool = True, index: bool = True, mode: str = "w") -> None:
        """
        def to_excel(header: bool = True, index: bool = True, mode: str = "w") -> None

        header
            Affiche les entetes de colonnes
            valeur par defaut = True
        index
            Affiche les entetes de lignes
            valeur par defaut = True
        mode
            a : ajoute les feuilles au fichier existant
            w : reinitialise les informations du fichier existant
        """
        pane_row: int = 1 if header else 0
        pane_col: int = 1 if index else 0
        with pd.ExcelWriter(self.__filename,
          date_format="DD-MM-YYYY",
          datetime_format="DD-MM-YYYY HH:MM:SS") as writer:  

            for sheet in self.__onglets:
                sheet.getData().to_excel(
                    writer, 
                    sheet_name=sheet.getName(),
                    freeze_panes=(pane_row, pane_col),
                    header=header,
                    index=index)


def ex2():
    sh = Onglet("Personnes")
    sh.add_colonnes(["Nom", "Age", "Nb", "Naissance"], ["", 0, 0, pd.Timestamp("20230101")])

    sh.add_line(["Christophe", 52, 1, pd.Timestamp("20230928")])
    sh.add_line(["Brigitte", 72, 2, pd.Timestamp("20230927")])

    sh.add_colonne("Suite", 0)
    sh.add_line(["JC", 69, 3, pd.Timestamp("20230926"), 5])

    sh.del_colonne("Age")
    sh.del_line(1)

    sh.add_line(["Romeo", 4, pd.Timestamp("20230925"), 6])

    sh.update_line(1, Nom="Kris", Nb=1, Suite=10)
    sh.rename_colonne("Nb", "Nombre")

    sh.sort(["Nombre"])

    sh.add_line(["Stephan", 8, pd.Timestamp("20230924"), 3])
    sh.add_line(["Romeo", 4, pd.Timestamp("20230923"), 6])

    sh.drop_duplicates()
    sh.sort_index()
    sh.reindex()
    sh.getData().describe()

    plus = Onglet("AutrePersonnes")
    plus.add_colonnes(["Nom", "Nombre", "Naissance", "Suite"], ["", 0, pd.Timestamp("20000101"), 0])
    plus.add_line(["Suite1", 1, pd.Timestamp("20230922"), 2])
    plus.add_line(["Suite2", 3, pd.Timestamp("20230921"), 4])

    somme = sh + plus
    somme.add_line(["dernier", 9, pd.Timestamp("20230920"), 99])
    somme.rename("Total")

    wb = Classeur("test.xlsx")
    wb.add_sheet(sh)
    wb.add_sheet(plus)

    somme.add_colonne("NewDate", somme.getData()["Naissance"]+pd.DateOffset(days=-1))
    # somme.getData()["NewDate"] = pd.to_datetime(somme.getData()["NewDate"])
    somme.add_colonne("DayOfWeek", somme.getData()["NewDate"].dt.dayofweek)
    wb.add_sheet(somme)

    if False:
        wb.drop_sheet(plus)
        wb.dropSheetByIndex(1)
        wb.dropSheetByName("Total")

    print(wb.getSheet("Total").getData()[:3])
    # wb.to_excel()


def ex1():
    js: dict = {}
    js["Nom"] = ["Alain", "Sylvie", "Albert", "Alain", "kris", "Gibritte", "Alphonse"]
    js["Age"] = [10, 11, 12, 13, 13, 14, 15]
    js["Nombre"] = [1, 3, 4, 7, 6, 7, 9]
    vpd = pd.DataFrame(js)

    # print(vpd.info())
    # vpd.to_excel("test.xlsx", sheet_name="age", index=False)

    result = vpd.pivot_table(
            values="Nombre",
            index="Nom",
            columns="Age",
            aggfunc="sum",
            margins=True)[10]

    print(result[result.notna()])

    # print(vpd[vpd["Age"] == 10]["Nom"])
    # vpd[vpd["Age"] == 10].iloc[1, 0] = "Titi"
    vpd.loc[vpd["Age"] == 10, "Nom"] = "Jumeau"
    vpd.iloc[3, 1] = 20

    print(vpd.sort_values(by="Nom").reset_index())
    # print(dir(vpd.groupby(["Nom"])["Age"]))
    # print(vpd[(vpd["Nom"] == "Jumeau") & (vpd["Age"] == 10)])


def ex3():
    df = pd.DataFrame({'Nom': ["Aze", "Qsd", "Wxc"], 'Age': [4, 5, 6], 'Nb': [7, 8, 9]})
    df.loc[len(df)] = ["Qfgh", None, 12]    
    print(df)


def ex4():
    c = Classeur("monClasseur.xlsx")
    sh = Onglet("Personnes")
    sh.add_colonnes(["Nom", "Prenom", "Sexe", "DateNaissance", "Adresse"], 
        [pd.StringDtype(), "", "", pd.Timestamp("20230101"), list()])

    # [pd.Series([], dtype="string"), "", "", pd.Timestamp("20230101"), list()])

    print(sh.getData().dtypes)
    # sh.add_colonnes(["Nom", "Prenom", "Sexe", "DateNaissance"], [pd.StringDtype(), "", "", pd.Timestamp("20230101")])

    # sh.add_line(["JACQUES", "Christophe", "M", pd.Timestamp("19710902"), [13, "bis", "rue du champ rond", 45000, "ORLEANS"]])
    sh.add_line(["JACQUES", "Christophe", "M", pd.Timestamp("19710902"), [13, "bis", "rue du champ rond", 45000, "ORLEANS"]])

    sh.add_line(["BERNARD", "Brigitte", "F", pd.Timestamp("19510916"), [12, "", "Athena", 45000, "ORLEANS"]])
    sh.add_colonne("DayOfWeek", sh.getData()["DateNaissance"].dt.day_of_week)

    print(sh.getData().iloc[0])

    c.add_sheet(sh)
    print(c.getSheets("Personnes"))


ex4()
