import html
import re
from datetime import date
from typing import Any, Dict, List

import requests
import xml.etree.ElementTree as ET
import pandas as pd
from openpyxl.styles import Alignment

USER_EMAIL = "atulbp1@hotmail.com"
HEADERS = {"User-Agent": f"ArchaeologyResearchScript/1.0 (mailto:{USER_EMAIL})"}

START_YEAR = 1976
END_YEAR = date.today().year

REGION_PATTERN = re.compile(
    r"\b(?:oman|omani|arabia|arabian peninsula|united arab emirates|uae|"
    r"abu dhabi|dubai|sharjah)\b",
    re.IGNORECASE,
)
PERIOD_PATTERN = re.compile(
    r"\b(?:bronze[\s-]+age|umm[\s-]+(?:an|al)[\s-]+nar|wadi[\s-]+suq|"
    r"hafit|third millennium|3rd millennium)\b",
    re.IGNORECASE,
)

MATERIAL_KEYWORDS = {
    "Ceramics": ["ceramic", "pottery", "sherd", "vessel"],
    "Soft stone": ["soft stone", "chlorite", "steatite"],
    "Beads": ["bead", "necklace", "pendant"],
    "Metal objects": ["metal", "copper", "bronze", "slag", "alloy"],
    "Textile fibers": ["textile", "fiber", "fabric", "weave"],
    "Seals": ["seal", "stamp seal", "cylinder seal"],
    "Shell objects": ["shell", "mollusk", "marine shell"],
    "Art and symbolism": ["art", "symbol", "iconography", "engraving"],
    "Human remains": ["human remain", "skeleton", "bone", "burial", "interment", "dental", "teeth"],
    "Faunal remains": ["faunal", "animal bone", "zooarchaeology", "livestock"],
    "Pollens": ["pollen", "palynology"],
    "Phytoliths": ["phytolith"],
    "Mortuary structures": ["tomb", "grave", "cairn", "mortuary", "sepulchre"],
    "Architectural elements": ["architecture", "wall", "dwelling", "structure", "building"]
}

METHOD_KEYWORDS = {
    "Isotopic analysis": ["isotope", "isotopic", "strontium", "carbon-13", "nitrogen-15"],
    "Proteomic analysis": ["proteom", "protein", "collagen"],
    "Genome analysis": ["dna", "genom", "ancient dna", "adna"],
    "Dental Calculus Examination": ["dental calculus", "microdebris", "starches"],
    "Paleopathological analysis": ["pathology", "paleopathol", "lesion", "trauma", "disease"],
    "Bioarchaeological analysis": ["bioarchaeol", "osteol", "skeletal"],
    "Detailed archaeological analysis": ["excavation", "stratigraph", "typology", "assemblage"],
    "Survey": ["survey", "remote sensing", "gis", "satellite"]
}

AIM_KEYWORDS = {
    "Excavation results": ["excavation", "fieldwork", "trench", "stratigraphy"],
    "Statistical analysis": ["statistical", "regression", "multivariate", "pca"],
    "Comparative analysis": ["comparative", "comparison", "cross-regional"],
    "Overview of findings": ["overview", "synthesis", "review", "summary"]
}

def search_openalex(max_results: int = 150) -> List[Dict[str, Any]]:
    print(" -> Fetching results from OpenAlex...")
    query_str = "Bronze Age Oman archaeology"
    url = "https://api.openalex.org/works"
    params = {"search": query_str, "filter": f"publication_year:{START_YEAR}-{END_YEAR}", "per_page": min(max_results, 200), "mailto": USER_EMAIL}
    results = []
    try:
        response = requests.get(url, params=params, headers=HEADERS, timeout=15)
        if response.status_code == 200:
            for item in response.json().get("results", []):
                abstract = ""
                inv_abstract = item.get("abstract_inverted_index")
                if inv_abstract:
                    words = {pos: word for word, positions in inv_abstract.items() for pos in positions}
                    abstract = " ".join([words[i] for i in sorted(words.keys())])
                affil = ""
                authorships = item.get("authorships", [])
                if authorships and authorships[0].get("institutions"):
                    affil = authorships[0]["institutions"][0].get("display_name", "")
                results.append({
                    "title": item.get("title", ""), "year": item.get("publication_year", ""),
                    "journal": item.get("primary_location", {}).get("source", {}).get("display_name", "") if item.get("primary_location") and item.get("primary_location").get("source") else "",
                    "doi": item.get("doi", "").replace("https://doi.org/", "") if item.get("doi") else "",
                    "abstract": abstract,
                    "affiliation": affil,
                    "authors": "; ".join(
                        author.get("author", {}).get("display_name", "")
                        for author in authorships
                        if author.get("author", {}).get("display_name")
                    ),
                    "url": item.get("primary_location", {}).get("landing_page_url", "") if item.get("primary_location") else "",
                    "source": "OpenAlex",
                })
        else:
            print(f"    Warning: OpenAlex returned HTTP {response.status_code}.")
    except Exception as e:
        print(f"    Warning: OpenAlex fetch issue: {e}")
    return results

def search_crossref(max_results: int = 100) -> List[Dict[str, Any]]:
    print(" -> Fetching results from CrossRef...")
    url = "https://api.crossref.org/works"
    params = {"query": "Bronze Age Oman trade diet mobility", "filter": f"from-pub-date:{START_YEAR}-01-01,until-pub-date:{END_YEAR}-12-31", "rows": max_results}
    results = []
    try:
        response = requests.get(url, params=params, headers=HEADERS, timeout=15)
        if response.status_code == 200:
            for item in response.json().get("message", {}).get("items", []):
                pub_parts = item.get("published-print", {}).get("date-parts") or item.get("published-online", {}).get("date-parts")
                year = pub_parts[0][0] if pub_parts and len(pub_parts[0]) > 0 else ""

                # Safely extract affiliation without throwing IndexError on empty author lists
                affil = ""
                authors = item.get("author", [])
                if authors and len(authors) > 0:
                    affils = authors[0].get("affiliation", [])
                    if affils and len(affils) > 0:
                        affil = affils[0].get("name", "")

                results.append({
                    "title": item.get("title", [""])[0] if item.get("title") else "",
                    "year": year,
                    "journal": item.get("container-title", [""])[0] if item.get("container-title") else "",
                    "doi": item.get("doi", ""),
                    "abstract": item.get("abstract", ""),
                    "affiliation": affil,
                    "authors": "; ".join(
                        " ".join(part for part in (author.get("given"), author.get("family")) if part)
                        for author in authors
                    ),
                    "url": item.get("URL", ""),
                    "source": "Crossref",
                })
        else:
            print(f"    Warning: Crossref returned HTTP {response.status_code}.")
    except Exception as e:
        print(f"    Warning: CrossRef fetch issue: {e}")
    return results

def search_pubmed(max_results: int = 100) -> List[Dict[str, Any]]:
    print(" -> Fetching results from PubMed...")
    term = f'("Bronze Age" OR "Umm an-Nar" OR "Wadi Suq") AND ("Oman" OR "Arabia") AND ({START_YEAR}/01/01[PDAT] : {END_YEAR}/12/31[PDAT])'
    results = []
    try:
        resp = requests.get("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi", params={"db": "pubmed", "term": term, "retmode": "json", "retmax": max_results}, headers=HEADERS, timeout=15)
        if resp.status_code == 200:
            id_list = resp.json().get("esearchresult", {}).get("idlist", [])
            if id_list:
                fetch_resp = requests.get("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi", params={"db": "pubmed", "id": ",".join(id_list), "retmode": "xml"}, headers=HEADERS, timeout=20)
                if fetch_resp.status_code == 200:
                    root = ET.fromstring(fetch_resp.content)
                    for article in root.findall(".//PubmedArticle"):
                        title = article.findtext(".//ArticleTitle", default="")
                        journal = article.findtext(".//Journal/Title", default="")
                        pub_date = article.find(".//JournalIssue/PubDate")
                        year = pub_date.findtext("Year", default="") if pub_date is not None else ""
                        abstract = "".join([elem.text for elem in article.findall(".//AbstractText") if elem.text])
                        doi = next((id_elem.text for id_elem in article.findall(".//ArticleId") if id_elem.get("IdType") == "doi"), "")
                        affil = article.findtext(".//Author/AffiliationInfo/Affiliation", default="")
                        authors = []
                        for author in article.findall(".//Author"):
                            name = " ".join(filter(None, [
                                author.findtext("ForeName"),
                                author.findtext("LastName"),
                                author.findtext("CollectiveName"),
                            ]))
                            if name:
                                authors.append(name)
                        pmid = article.findtext(".//PMID", default="")
                        results.append({
                            "title": title, "year": year, "journal": journal, "doi": doi,
                            "abstract": abstract, "affiliation": affil,
                            "authors": "; ".join(authors),
                            "url": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/" if pmid else "",
                            "source": "PubMed",
                        })
                else:
                    print(f"    Warning: PubMed fetch returned HTTP {fetch_resp.status_code}.")
        else:
            print(f"    Warning: PubMed search returned HTTP {resp.status_code}.")
    except Exception as e:
        print(f"    Warning: PubMed fetch issue: {e}")
    return results

def classify_text(text: str, keyword_dict: Dict[str, List[str]]) -> str:
    if not text: return ""
    text_lower = text.lower()
    matches = [cat for cat, kws in keyword_dict.items() if any(re.search(r'\b' + re.escape(kw) + r'\b', text_lower) for kw in kws)]
    return "; ".join(matches) if matches else ""

def determine_methodology(method_used: str) -> str:
    if not method_used:
        return ""
    quant_terms = ["Isotopic", "Proteomic", "Genome", "Statistical", "Survey"]
    has_quant = any(term in method_used for term in quant_terms)
    has_qual = "Detailed archaeological" in method_used or "Review" in method_used
    return "Mixed" if (has_quant and has_qual) else ("Quantitative" if has_quant else "Qualitative")

def clean_abstract(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = re.sub(r"<[^>]*>", " ", html.unescape(text))
    text = re.sub(r"[\u2010-\u2015\u2212]", "-", text)
    return re.sub(r"\s+", " ", text).strip()

def is_relevant(paper: Dict[str, Any]) -> bool:
    title = clean_abstract(paper.get("title", ""))
    if not title:
        return False
    text = f"{title} {clean_abstract(paper.get('abstract', ''))}"
    return bool(REGION_PATTERN.search(text) and PERIOD_PATTERN.search(text))

def main():
    raw_results = search_openalex() + search_crossref() + search_pubmed()
    seen_dois = set()
    seen_titles = set()
    deduped = []
    for item in raw_results:
        doi = item.get("doi", "").lower().strip()
        title = re.sub(r"\W+", "", item.get("title", "").lower())
        if (
            title
            and (not doi or doi not in seen_dois)
            and title not in seen_titles
            and is_relevant(item)
        ):
            if doi:
                seen_dois.add(doi)
            seen_titles.add(title)
            deduped.append(item)

    if not deduped:
        raise RuntimeError(
            "No relevant publications were returned. Check the source warnings or broaden the search."
        )

    rows = []
    for paper in deduped:
        abstract = clean_abstract(paper.get("abstract", ""))
        text = f"{paper['title']} {abstract}"
        mat = classify_text(text, MATERIAL_KEYWORDS)
        aim = classify_text(text, AIM_KEYWORDS)
        method_used = classify_text(text, METHOD_KEYWORDS)
        publication_year = paper.get("year")
        try:
            publication_year = int(publication_year)
        except (TypeError, ValueError):
            publication_year = ""

        rows.append({
            "Title": paper["title"],
            "Authors": paper.get("authors", ""),
            "Publication year": publication_year,
            "Journal": paper["journal"],
            "DOI": paper["doi"],
            "URL": paper.get("url") or (
                f"https://doi.org/{paper['doi']}" if paper.get("doi") else ""
            ),
            "Abstract": abstract,
            "Cultural material (keyword-coded)": mat,
            "Aim (keyword-coded)": aim,
            "Methodology": determine_methodology(method_used),
            "Method used (keyword-coded)": method_used,
            "Primary affiliation of primary author": paper["affiliation"],
            "Source": paper.get("source", ""),
        })

    df = pd.DataFrame(rows).sort_values(
        ["Publication year", "Title"], ascending=[False, True], na_position="last"
    )
    output_path = "Search_Results_Populated.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data information sheet")
        worksheet = writer.sheets["Data information sheet"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        for column_cells in worksheet.columns:
            header = column_cells[0].value
            width = min(max(len(str(cell.value or "")) for cell in column_cells[:30]) + 2, 42)
            worksheet.column_dimensions[column_cells[0].column_letter].width = max(width, 14)
            if header == "Abstract":
                for cell in column_cells[1:]:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

    print(f"Retained {len(df)} relevant publications from {len(raw_results)} source records.")
    print(f"Done! Data written to {output_path}")

if __name__ == "__main__":
    main()
