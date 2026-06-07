<p align="center">
  <img src="GitHub.png" width="150">
</p>

# Matris
Matris is a Python-based software tool for automated analysis of ethological data collected during observations of maternal and social behavior.

> A Python tool for analyzing maternal and social behavior protocols

**Matris** is a Python-based software pipeline that automates the processing of ethological data.  
The project was developed as part of a thesis and is intended for animal behavior researchers.

---

## 📌 Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Input Data Format](#input-data-format)
- [Output Data Format](#output-data-format)
- [Authors](#authors)
- [License](#license)
- [Citation](#citation)
- [Keywords](Keywords)
- [Installation and Quick Star](#Installation-and-Quick-Star)

---

## 🎯 Features

| Feature | Description |
|---------|-------------|
| **Frequency analysis** | Counts behavioral elements by actors and groups |
| **Duration analysis** | Calculates feeding time and social interaction duration |
| **Normalization** | Adjusts values for comparative analysis |
| **Aggregation** | Generates a summary table ready for statistical processing |

---

## 💻 Requirements

- Python 3.8+
- pandas ≥ 2.0.0
- numpy ≥ 1.24.0
- openpyxl (for Excel support)

---

## 📂 Input Data Format
The protocol file (CSV or Excel) must contain the following columns:

|Column |Type |	Description |
|---------|-------------|----------------|
|timestamp |	datetime |	Time of the event |
|actor |	str |	Individual performing the behavior |
|behavior |	str |	Behavioral element |
|duration |	float |	Duration (if applicable) |
---

## 📊 Output Data Format
The summary table (CSV/Excel) contains:

|Group |	Number of actions |	Duration of actions (s) |	Normalized frequency |
|------|--------------------|-------------------------|----------------------|
|group of animals |	45 |	1024.5  |	56.2 |
|a specific animal |	12 |	145.2 |	21.8 |
|pairs of animals |	8 |	69.6 |	10.9 |

---

## 👥 Authors
|Name |	Affiliation |
|-----|-------------|
|Khrisiya F. Yagushova | Russian State Agrarian University — Moscow Timiryazev Agricultural Academy
|Alexandra S. Fetisova | Severtsov Institute of Ecology and Evolution of the Russian Academy of Sciences
|Mariya N. Erofeeva | Severtsov Institute of Ecology and Evolution of the Russian Academy of Sciences
|Sergey V. Naidenko | Severtsov Institute of Ecology and Evolution of the Russian Academy of Sciences

---

## 📄 License
This project is distributed under the MIT License.
See the LICENSE file for details.

---

## 📖 Citation
If you use Matris in your research, please cite:

Yagushova K.F., Fetisova A.S., Erofeeva M.N., Naidenko S.V. E. (2026). Matris: A tool for analyzing maternal and social behavior protocols.

A CITATION.cff file is available in the repository root.

---

## 📬 Contact
For collaboration inquiries: hrisiayagushova@yandex.ru

---

## Keywords
Matris, ethology, behavioral analysis, maternal behavior, social behavior, behavioral frequencies, duration analysis, data normalization, summary table, Python, pandas, open-science, reproducible research, animal behavior, behavioral ecology, observation protocols

---

## 📦 Installation and Quick Start

```bash
# Clone the repository
git clone https://github.com/your-username/matris.git
cd matris

# Install dependencies
pip install -r requirements.txt

---

Quick Start

from matris import Matris

# Initialize
analyzer = Matris()

# Load data
data = analyzer.load("data/raw/observations.csv")

# Calculate frequencies
freq = analyzer.frequency(data, group_by="actor")

# Calculate durations
dur = analyzer.duration(data, behavior="feeding")

# Normalize
norm = analyzer.normalize(data, baseline="total_time")

# Generate summary
summary = analyzer.summary(data)

print(summary)
