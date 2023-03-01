
import json
import openpyxl
from pathlib import Path
from itertools import islice
from SPARQLWrapper import SPARQLWrapper, JSON

endpoint = SPARQLWrapper("https://skynet.coypu.org/coypu-internal/")
endpoint.setCredentials(user="katherine", passwd = "0CyivAlseo")

prefixes = "PREFIX coy: <https://schema.coypu.org/global#> " + \
           "PREFIX xsd: <http://www.w3.org/2001/XMLSchema#> " + \
           "PREFIX rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#> "  + \
           "PREFIX rdfs: <http://www.w3.org/2000/01/rdf-schema#> "  + \
           "PREFIX data: <https://data.coypu.org/> " + \
           "PREFIX skos: <http://www.w3.org/2004/02/skos/core#> " + \
           "PREFIX geo: <http://www.opengis.net/ont/geosparql#> " + \
           "PREFIX wpi: <https://schema.coypu.org/world-port-index#> " + \
           "PREFIX wdt: <http://www.wikidata.org/prop/direct/> " + \
           "PREFIX wd: <http://www.wikidata.org/entity/> " + \
           "PREFIX em: <https://schema.coypu.org/em-dat#> " + \
           "PREFIX schema: <http://schema.org/> " + \
           "PREFIX gta:  <https://schema.coypu.org/gta#>"

def main():
    ds_file = Path.cwd().joinpath("QA Questions CoyPu.xlsx")
    wb_obj = openpyxl.load_workbook(ds_file)
    active_sheet = wb_obj.active

    ds_questions = {}
    ds_questions["dataset"] = {"id": "coypu-kg-questions-all"}
    ds_questions["questions"] = []

    for idx, row in enumerate(islice(active_sheet.iter_rows(max_col=4, values_only=True), 1, None)):
        en_str, de_str, sparql, _ = row
        if en_str is None and de_str is None and sparql is None:
            break

        question = {}
        question["id"] = str(idx)
        question["answers"] = []

        q = []
        q_en, q_de = {}, {}
        q_en["string"] = en_str
        q_en["language"] = "en"
        q.append(q_en)
        q_de["string"] = de_str
        q_de["language"] = "de"
        q.append(q_de)
        question["question"] = q

        query = {}
        query["sparql"] = prefixes + sparql
        endpoint.setQuery(query["sparql"])
        endpoint.setReturnFormat(JSON)
        result = endpoint.query().convert()

        question["answers"].append(result)
        question["query"] = query
        ds_questions["questions"].append(question)

    with open("coypu-questions-all.json", "w") as outfile:
        json.dump(ds_questions, outfile, indent=4)


if __name__ == "__main__":
    main()
