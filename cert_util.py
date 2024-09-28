import re
import pandas as pd


# name = "Mr. Satish Annigeri"
# s = re.sub(r"^Mr. ", "", name)
# print(name, s)

# name = "Ms. Uma Annigeri"
# s = re.sub(r"^Ms. ", "", name)
# print(name, s)

# name = "Mrs. Uma Annigeri"
# s = re.sub(r"^Mrs. ", "", name)
# print(name, s)


def name_salutation(
    name: str, gender: str, strip_pattern: list[str], salut: str = ""
) -> str:
    for pattern in strip_pattern:
        if name.startswith(pattern):
            bare_name = re.sub(pattern, "", name, count=1)
            bare_name = bare_name.strip()
        else:
            bare_name = name

    match gender.upper()[0]:
        case "M":
            return f"Mr. {bare_name}"
        case "F":
            return f"Ms. {bare_name}"
        case _:
            return f"{salut} {bare_name}" if salut else bare_name


def mod_name(bare_name: str, gender: str) -> str:
    match gender[0]:
        case "M":
            return f"Mr. {bare_name}"
        case "F":
            return f"Ms. {bare_name}"
        case _:
            return bare_name


if __name__ == "__main__":
    strip_pattern = ["Mr.", "Ms.", "Mrs.", "Dr."]
    name = "Mr. Satish Annigeri"
    print(name_salutation(name, "m", strip_pattern))
    print(name_salutation(name, "f", strip_pattern))
    name = "Dr.Satish Annigeri"
    print(name_salutation(name, "M", strip_pattern))
    print(name_salutation(name, "X", strip_pattern))
    print(name_salutation(name, "X", strip_pattern, "Dr."))

    df = pd.read_excel("test01.xlsx")
    df["gender"] = df["gender"].str[0].str.upper()
    df["bare_name"] = df["name"].str.replace(
        r"^[M]{1}[rs]{1}[s]*\.[ ]*", "", regex=True
    )
    df["slaut_name"] = df.apply(lambda r: mod_name(r["bare_name"], r["gender"]), axis=1)
    print(df)
