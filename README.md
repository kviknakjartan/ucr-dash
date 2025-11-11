# Climate Change in Graphs

A data-visualization web app built with [Streamlit](https://streamlit.io) to explore how climate change is affecting our planet—through temperature anomalies, greenhouse-gas emissions, economic indicators and more.

## 🎯 Project Overview  
This app brings together multiple datasets (e.g., temperature change, CO₂ / N₂O / CH₄ emissions) and lets users interactively explore global and country-level trends.  
Key objectives:  
- Make climate data accessible and visually engaging  
- Allow users to compare countries, indicators and time-periods  
- Highlight long-term patterns and correlations 

## 🚀 Features  
- Interactive dashboards powered by Streamlit  
- Multiple graph types (line charts, bar charts, choropleth maps, scatter plots)  
- Country selector and time-range filters  
- Responsive layout so viewable on desktop or mobile  

## 🛠 Tech Stack  
- Python (≥ 3.7)  
- Streamlit  
- Pandas  
- Plotly (or Plotly Express) for interactive charts  
- GeoJSON / Geopandas for map-visualizations (if used)  
- Optionally caching via `@st.cache` to speed up loading of large datasets  

## 📂 Project Structure  
├── data/ ← raw & processed data files
├── pages/ ← pages in the app except Home
│ ├── Temperature.py ← Global temperature and greenhouse gas concentrations
│ ├── Energy.py ← World energey production and consumption
│ ├── Emissions.py ← Ice sheets, snow cover, sea ice extent
│ ├── Ice.py ← Ice sheets, snow cover, sea ice extent
│ ├── Maps.py ← Various global spatial distributions of climate indicators and effects
│ ├── Ocean.py ← Sea level rise, acidity, ocean heat content
│ └── Quantities.py ← Physical quantities such as climate sensitivity and radiative forcing
├── Home.py ← Streamlit entry-point
├── get_data.py ← Module for loading and handling of data
├── requirements.txt ← Python dependencies
├── LICENSE ← MIT license file
└── README.md ← this file


## 📥 Installation & Usage  
1. Clone the repository:  
   ```bash
   git clone https://github.com/YourUsername/your-repo.git
   cd your-repo
2.Install dependencies:
   pip install -r requirements.txt
3.Run the app:
   streamlit run src/app.py
4.Open the URL printed in your terminal (typically http://localhost:8501) in your browser.

🧮 Data Sources
See references on each page.

📄 License

MIT License
Copyright (c) 2025 Kjartan Pétursson

📞 Contact

For further questions or comments:

Email: kjartanbrjann@gmail.com

GitHub: https://github.com/kviknakjartan

Thank you for using “Climate Change in Graphs” — we hope it helps you gain deeper insights into how our planet is changing and what that means for humanity and ecosystems.