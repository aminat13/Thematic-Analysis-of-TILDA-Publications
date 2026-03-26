# Thematic-Analysis-of-TILDA-Publications

This project explores a workflow for organising and synthesising publication metadata from TILDA (The Irish Longitudinal Study on Ageing) into broader research themes.

The workflow was designed to:

- reorganise papers by topic
- identify relationships between topics using co-occurrence analysis
- group related topics into broader thematic communities
- retrieve abstracts automatically using DOI-based lookup
- support structured, theme-level synthesis suitable for research communication

This repository focuses on the analysis pipeline, not the underlying proprietary dataset.

## Project Overview

The workflow proceeds in five main stages:

### Data preparation
- Load publication metadata from Excel
- Parse topic IDs associated with each paper
- Standardise topic structure for analysis
### Topic relationship analysis
- Generate topic co-occurrence pairs within each paper
- Count how often topic pairs appear together
- Filter for stronger topic relationships
### Theme detection
- Build a weighted topic network from strong co-occurrence pairs
- Apply community detection to identify broader thematic clusters
- Map topic IDs back to human-readable topic names
### Abstract retrieval
- Retrieve abstracts automatically from DOI metadata
- Use Crossref first, with PubMed as a fallback
- Store retrieval status and source alongside paper metadata
### Theme-level synthesis
Summarise paper abstracts into structured fields:
- study aim
- main finding
- methods used
- limitation
- evidence sentence
Combine paper-level summaries into broader theme-level syntheses

## Example Outputs

The workflow produces outputs such as:

- topic-level Excel sheets
- co-occurrence tables
- weighted topic network graphs
- thematic communities
- abstract retrieval tables
- theme-level narrative summaries

## Methods Used

This project uses:

- pandas for spreadsheet wrangling
- itertools and collections.Counter for topic pair generation
- networkx for graph construction and community detection
- requests and xml.etree.ElementTree for DOI-based abstract retrieval
- matplotlib for topic network visualisation

## Data Availability

- Due to data ownership and confidentiality, the original dataset used in this project is not included.
- This repository focuses on the analysis pipeline and methodology. 
