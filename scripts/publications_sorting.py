#loading excel sheet and viewing column names
import pandas as pd

df = pd.read_excel(example_publications.xlsx, sheet_name="Publications")
print(df.columns)

df["Topics"] = df["Topics"].str.split() #splitting by spaces
print(df['Topics'].head())

#creating for loop to count pairs in topic list
from itertools import combinations #package that makes combos
from collections import Counter #package that counts combos

pairs = Counter()

for topics in df["Topics"]:
    unique_topics = list(set(topics)) #remove duplicates within a pair
    for pair in combinations(unique_topics, 2): #creating pairs from topics in row
        pair = tuple(sorted(pair)) #forcing a consistent order
        pairs[pair] += 1 #adding to pair count

print(len(pairs)) #viewing how many pairs

#creating dataframe
pairs_df = pd.DataFrame([(topic1, topic2, count) for (topic1, topic2), count in pairs.items()], #unpacking pairs in counter
    columns=["Topic_1", "Topic_2", "Count"])

#sorting from the highest frequency to the lowest
pairs_df = pairs_df.sort_values("Count", ascending=False)

#performing sanity checks
print("Unique pair types:", len(pairs_df))
print("Total pair occurrences:", pairs_df["Count"].sum())
print("Top 10 pairs:")
print(pairs_df.head(10))

#creating new Excel file
with pd.ExcelWriter("example_topics_output", mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    pairs_df.to_excel(writer, sheet_name="Topic_Pairs", index=False)

#creating dataframe of most frequent pairs
strong_pairs = pairs_df[pairs_df["Count"] >= 10]
strong_pairs = strong_pairs.sort_values("Count", ascending=False)
print(len(strong_pairs))
print(strong_pairs.head(5))

#performing graph clustering
import networkx as nx #package for working with network graphs
from networkx.algorithms import community #package to find clusters in graphs

G = nx.Graph() #nodes=topics, edges=connections btw topics

for index, row in strong_pairs.iterrows(): #looping through rows in dataframe
    G.add_edge(row["Topic_1"], row["Topic_2"], weight=row["Count"]) #connecting topic 1 and 2

communities = community.greedy_modularity_communities(G, weight="weight") #finds clusters with low WSS and high BSS

for i, comm in enumerate(communities, 1): #gives numbering to communities (list of sets)
    print(f"Community {i}: {sorted(comm)}")

#creating dictionary for topic id and names
topic_df = pd.read_excel("example_publications.xlsx", sheet_name="Keywords")
topic_map = dict(zip(topic_df["id"], topic_df["keyword"]))

#creating communities with topic names
for i, comm in enumerate(communities, 1):
    topic_names = [topic_map.get(int(topic), "Unknown") for topic in sorted(comm)]
    print(f"Community {i}: {topic_names}")

#visualising network graph
import matplotlib.pyplot as plt
import networkx as nx
from matplotlib.patches import Patch

#creating labels
labels = {node: topic_map.get(int(node), "Unknown") for node in G.nodes()}

#assigning a color to each node based on its community
tilda_palette = [
    "#4A90E2",
    "#7ED321",
    "#50E3C2",
    "#B8E986"
]

color_map = {}
for i, comm in enumerate(communities):
    color = tilda_palette[i % len(tilda_palette)]
    for node in comm:
        color_map[node] = color
node_colors = [color_map[node] for node in G.nodes()]

#creating legend based on colour map
community_names = [
    "Physical Health & Ageing",
    "Social & Economic Wellbeing",
    "Mental Health",
    "Healthcare & Service Use"
]

legend_elements = []
for i, comm in enumerate(communities):
    color = tilda_palette[i % len(tilda_palette)]
    label = community_names[i] if i < len(community_names) else f"Community {i+1}"

    legend_elements.append(
        Patch(facecolor=color, edgecolor='black', label=label)
    )
plt.legend(handles=legend_elements, loc='best', fontsize=9)

plt.figure(figsize=(13, 10))
pos = nx.spring_layout(G, seed=42, k=0.7)
edges = G.edges(data=True)
weights = [d["weight"] for (_, _, d) in edges]
nx.draw(
    G,
    pos,
    labels=labels,
    node_color=node_colors,
    node_shape="s",          # square nodes
    node_size=1400,
    font_size=8,
    font_color="#1F1F1F",    # dark grey labels
    edge_color="#B7C9D6",    # soft grey-blue edges
    width=[w / 20 for w in weights]
)
fig.savefig(
    "working_directory",
    dpi=300,
    bbox_inches='tight',
    pad_inches=0.3
)
plt.show()

#finding topics not included in communities using list comprehension: t=topic ID, d=number of connections
all_topics = set(t for topics in df["Topics"] for t in topics)
graph_topics = set(G.nodes())
missing_topics = all_topics - graph_topics
print(len(missing_topics))
print([(t, topic_map.get(int(t), "Unknown")) for t in missing_topics])

#finding topics with high number of connections
degree = dict(G.degree()) #calculates how many connections a node (topic) has
sorted_degree = sorted(degree.items(), key=lambda x: x[1], reverse=True) #sorting topics by degree (most connected first)
[(t, topic_map.get(int(t), "Unknown"), d) for t, d in sorted_degree[:10]] #creating list of 10 most connected topics

#separating topic column
working_df = df.copy()
working_df = working_df.explode("Topics")
working_df["Topics"] = working_df["Topics"].str.strip() #getting rid fo white space
print(working_df['Topics'].head())

#adding topics as sheets to Excel file
with pd.ExcelWriter("example_topics_output.xlsx", engine="xlsxwriter") as writer: #creating new excel file
    for topic, group_df in working_df.groupby("Topics"):
        topic_name = topic_map.get(int(topic), "Unknown") #setting topic and integer and return unknown if no corresponding name found
        sheet_name = f"{topic}_{topic_name}"[:31] #limiting sheet name to 31 characters
        group_df.to_excel(writer, sheet_name=sheet_name, index=False)

#viewing sheet names
xls = pd.ExcelFile(r"example_topics_output.xlsx")
print(xls.sheet_names)

#defining functions to receive abstracts
import requests
import re
import xml.etree.ElementTree as ET

#defining crossref function
def get_abstract_crossref(doi):
    url = f"https://api.crossref.org/works/{doi}"
    try: #fail safe to prevent crashing
        response = requests.get(url, timeout=20)
        response.raise_for_status() #if error go to except
        data = response.json()["message"] #storing info in dictionary
        abstract = data.get("abstract", None) #retrieving abstract from dictionary
        if abstract: #checking if abstract was actually found
            clean_abstract = re.sub(r"<.*?>", "", abstract) #removing xml format
            return clean_abstract.strip(), "Crossref", "Found" #returning abstract text, source name, status
        else:
            return None, "Crossref", "Not found"
    except Exception: #fail safe to prevent crashing
        return None, "Crossref", "Error"

#defining pubmed function
def get_abstract_pubmed(doi):
    try:
        #getting pmid from doi
        search_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?db=pubmed&term={doi}[DOI]&retmode=json"
        search_response = requests.get(search_url, timeout=20)
        search_response.raise_for_status() #if error go to except
        search_data = search_response.json() #storing info in dictionary
        id_list = search_data.get("esearchresult", {}).get("idlist", []) #getting pmid from doi
        if not id_list:
            return None, "PubMed", "Not found"
        pmid = id_list[0]
        fetch_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&id={pmid}&retmode=xml"
        fetch_response = requests.get(fetch_url, timeout=20)
        fetch_response.raise_for_status() #if error go to except
        root = ET.fromstring(fetch_response.content)
        abstract_elements = root.findall(".//AbstractText")
        if abstract_elements:
            abstract = " ".join([elem.text for elem in abstract_elements if elem.text]) #retrieving abstract from list and cleaning it
            return abstract.strip(), "PubMed", "Found" #returning abstract text, source name, status
        return None, "PubMed", "Not found"
    except Exception: #fail safe to prevent crashing
        return None, "PubMed", "Error"

#defining master function
def get_abstract(doi):
    abstract, source, status = get_abstract_crossref(doi)
    if abstract:
        return abstract, source, status
    abstract, source, status = get_abstract_pubmed(doi)
    if abstract:
        return abstract, source, status
    return None, "None", "Not found"

#subsetting topic from df and adding new columns for storage (change df name!!!)
topic67_df = working_df[working_df["Topics"] == "67"].copy()
topic67_df["Abstract"] = None
topic67_df["Abstract_Source"] = None
topic67_df["Retrieval_Status"] = None

#using function on every row
results = []

for index, row in topic67_df.iterrows():
    doi = row["doi"]
    abstract, source, status = get_abstract(doi)
    results.append({
        **row.to_dict(),  #dictionary unpacking
        "Abstract": abstract,
        "Abstract_Source": source,
        "Retrieval_Status": status
    })
topic67_abstracts_df = pd.DataFrame(results)
topic67_abstracts_df.to_excel("example_topic_abstracts.xlsx", index=False)

#checking status
topic67_abstracts_df["Retrieval_Status"].value_counts()

#adding to same file with another sheet (change df and sheet name !!!)
with pd.ExcelWriter("example_topic_abstracts.xlsx", mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    topic67_abstracts_df.to_excel(writer, sheet_name="Topic_67", index=False)
