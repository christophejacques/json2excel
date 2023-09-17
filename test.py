import pandas as pd
from typing import Any, Self


def definition():
    [print(f"{attrib:20}", end="") for attrib in filter(lambda x: not x.startswith("_"), 
        dir(pd.DataFrame))]
    print()
    # help(pd.DataFrame.drop)
    exit()


# definition()


class Sheet:

    def __init__(self, nom_colonnes: list[str]):
        if not isinstance(nom_colonnes, list) and not isinstance(nom_colonnes, tuple):
            raise TypeError(f"Le type des donnees doit etre 'list', au lieu de '{nom_colonnes.__class__.__name__}'.")

        self.__nom_col: list[str] = nom_colonnes
        self.__nb_cols = len(nom_colonnes)
        self.__data: pd.DataFrame = pd.DataFrame()
        self.__last_index: int = 0

    def add_colonne(self, nom_colonne: str, valeur_defaut: Any) -> None:
        if nom_colonne in self.__nom_col:
            raise ValueError(f"La colonne '{nom_colonne}' existe deja.")

        self.__data[nom_colonne] = valeur_defaut
        self.__nom_col.append(nom_colonne)
        self.__nb_cols += 1

    def rename_colonne(self, old_name: str, new_name: str) -> None:
        self.__data.rename(columns={old_name: new_name}, inplace=True)
        self.__nom_col[self.__nom_col.index(old_name)] = new_name

    def del_colonne(self, colonne_name: str) -> None:
        # self.__data = self.__data.drop(columns=[colonne_name])
        self.__data.drop(columns=[colonne_name], inplace=True)
        self.__nom_col.remove(colonne_name)
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
        df: pd.DataFrame = pd.DataFrame([valeurs], columns=self.__nom_col, index=[self.__last_index])
        self.__data = pd.concat([self.__data, df])

    def del_line(self, line_idx: int) -> None:
        self.__data.drop(index=[line_idx], inplace=True)
        self.__last_index = self.__data.index.max() + 1

    def update_line(self, index: int, **colonnes) -> None:
        for idx, colonne in enumerate(colonnes):
            # print(f"{colonne} set to: {colonnes[colonne]}")
            # self.__data[colonne].iloc[index-1] = colonnes[colonne]
            self.__data.iloc[index-1, idx] = colonnes[colonne]

    def drop_duplicates(self) -> None:
        self.__data.drop_duplicates(inplace=True)

    def getData(self) -> pd.DataFrame:
        return self.__data

    def to_excel(self, nom_fichier: str) -> None:
        self.__data.to_excel(nom_fichier)

    def __add__(self, other): 
        if "|".join(self.__nom_col) != "|".join(other.__nom_col):
            raise ValueError("Le nom des colonnes sont differents")

        nombre1 = 1+self.__data.index.shape[0]
        data1 = pd.DataFrame(self.__data.values, 
            columns=self.__data.columns, 
            index=range(1, nombre1))

        nombre2 = other.__data.index.shape[0]
        data2 = pd.DataFrame(other.__data.values, 
            columns=other.__data.columns, 
            index=range(nombre1, nombre1+nombre2))

        sheet: Sheet = Sheet(self.__nom_col)
        sheet.__data = pd.concat([data1, data2])
        sheet.__last_index = sheet.__data.index.shape[0]

        return sheet
        # return pd.concat([self.__data, other.__data])

    def __str__(self) -> str:
        return f"{self.__data}"


sh = Sheet(["Nom", "Age", "Nb"])
sh.add_line(["Christophe", 52, 1])
sh.add_line(["Brigitte", 72, 2])

sh.add_colonne("Suite", 0)
sh.add_line(["JC", 69, 3, 5])
print(sh.getData())

sh.del_colonne("Age")
sh.del_line(1)

sh.add_line(["Romeo", 4, 6])

print(sh.getData())

sh.update_line(1, Nom="Kris", Nb=1, suite=1)
sh.rename_colonne("Nb", "Nombre")

sh.sort(["Nombre"])
print("Sort by Nombre")
print(sh.getData())

sh.add_line(["Stephan", 8, 3])
sh.add_line(["Romeo", 4, 6])

sh.drop_duplicates()
print()
sh.sort_index()
sh.reindex()
print(sh.getData())

plus = Sheet(["Nom", "Nombre", "Suite"])
plus.add_line(["Suite1", 1, 2])
plus.add_line(["Suite2", 3, 4])
print(plus.getData())

print()
wb = sh + plus
print(wb)

print()
wb.add_line(["dernier", 9, 99])
print(wb)

# wb.to_excel("test.xlsx")


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
    print()
    print(vpd.sort_values(by="Nom"))
    print()
    print(vpd)
    # print()
    # print(dir(vpd.groupby(["Nom"])["Age"]))
    # print(vpd[(vpd["Nom"] == "Jumeau") & (vpd["Age"] == 10)])


def ex3():
    df = pd.DataFrame({'Nom': ["Aze", "Qsd", "Wxc"], 'Age': [4, 5, 6], 'Nb': [7, 8, 9]})
    df.loc[len(df)] = ["Qfgh", None, 12]    
    print(df)


# ex1()
