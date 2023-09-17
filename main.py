import pandas as pd
import json


def pandas_read_json():
    js = pd.read_json("data.json")
    print(js)


def open_json():
    with open("data.json") as f:
        js = json.load(f)
        # print(js)
        with pd.ExcelWriter('output.xlsx') as writer:  
            for sheet in js:
                vpd = pd.DataFrame({sheet: js[sheet]})
                vpd.to_excel(writer, sheet_name=sheet)

            # print(writer.book)

            for sheet in writer.sheets:
                print("Sheet:", writer.sheets[sheet].title)
                maxsize: list = [0 for _ in range(writer.sheets[sheet].max_column)]
                for r in writer.sheets[sheet].rows:
                    for c in range(writer.sheets[sheet].max_column):
                        maxsize[c] = max(maxsize[c], len("" if r[c].value is None else str(r[c].value)))

                for r in writer.sheets[sheet].rows:
                    print(f"L{r[0].row}", end=" | ")
                    for c in range(writer.sheets[sheet].max_column):
                        chaine = " " * maxsize[c]
                        if r[c].value is not None:
                            chaine = f"{str(r[c].value)}{chaine}"
                        print(chaine[:maxsize[c]], end=" | ")
                    print()
                    if r[0].row == 1:
                        print("-" * (10 + sum(maxsize)))
                print()


open_json()
