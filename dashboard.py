import sys
from pathlib import Path
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

try:
    import streamlit as st
except ImportError:
    print("Streamlit is required. Install with: pip install streamlit")
    sys.exit(1)

# Configuration
DATA_FOLDER = Path(__file__).resolve().parent
FILE_PATTERNS = (
    "Bhada Oscillation Data_*.xlsx",
    "Bhadla Oscillation Data_*.xlsx",
    "Bhada Oscillation Data_*.xls",
    "Bhadla Oscillation Data_*.xls",
    "Bhada Oscillation Data_*.csv",
    "Bhadla Oscillation Data_*.csv",
)

def _discover_files() -> list[Path]:
    files: list[Path] = []
    for pattern in FILE_PATTERNS:
        files.extend(DATA_FOLDER.glob(pattern))
    return sorted(set(files))

def _load_one_file(path: Path) -> pd.DataFrame | None:
    try:
        if path.suffix.lower() in {".xlsx", ".xls"}:
            df = pd.read_excel(path)
        elif path.suffix.lower() == ".csv":
            df = pd.read_csv(path)
        else:
            return None

        # Strip whitespace from column names
        df.columns = df.columns.str.strip()
        
        # Validate required columns
        required = ["STARTDATE", "HZ", "VPM"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            print(f"Skipping {path.name}: missing columns {missing}")
            return None

        # 1. Convert timestamp to datetime
        df["STARTDATE"] = pd.to_datetime(df["STARTDATE"], errors="coerce", dayfirst=True)
        df = df.dropna(subset=["STARTDATE"]).copy()
        if df.empty:
            print(f"Skipping {path.name}: no valid STARTDATE values")
            return None

        # 2. Coerce numeric columns
        df["HZ"] = pd.to_numeric(df["HZ"], errors="coerce")
        df["VPM"] = pd.to_numeric(df["VPM"], errors="coerce")
        
        # 3. FIX: Distribute samples within the same minute
        # Spreads high-frequency (20Hz) samples evenly across the 60 seconds of each minute.
        df['sample_idx'] = df.groupby('STARTDATE').cumcount()
        df['samples_in_min'] = df.groupby('STARTDATE')['STARTDATE'].transform('count')
        
        df['PRECISE_TIME'] = df['STARTDATE'] + pd.to_timedelta(
            (df['sample_idx'] / df['samples_in_min']) * 60, unit='s'
        )

        # 4. Data Cleaning - Handle missing values in HZ using linear interpolation
        df["HZ"] = df["HZ"].interpolate(method="linear", limit_direction="both")
        df["SOURCE_FILE"] = path.stem
        
        return df.drop(columns=['sample_idx', 'samples_in_min'])
    except Exception as e:
        print(f"Skipping {path.name}: {e}")
        return None

@st.cache_data(show_spinner=False)
def load_all_data() -> pd.DataFrame:
    files = _discover_files()
    if not files:
        return pd.DataFrame()
    
    loaded = []
    for f in files:
        df_tmp = _load_one_file(f)
        if df_tmp is not None and not df_tmp.empty:
            loaded.append(df_tmp)
    
    return pd.concat(loaded, ignore_index=True) if loaded else pd.DataFrame()

def create_oscillation_plot(df: pd.DataFrame):
    df_sorted = df.sort_values("PRECISE_TIME")

    fig = make_subplots(
        rows=2, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.08,
        subplot_titles=("Frequency (HZ)", "Voltage/Power Magnitude (VPM)")
    )

    # Top Plot: HZ
    fig.add_trace(
        go.Scatter(
            x=df_sorted["PRECISE_TIME"],
            y=df_sorted["HZ"],
            name="Frequency (HZ)",
            line=dict(color="#00D4FF", width=1.2),
            hovertemplate="Time: %{x}<br>Freq: %{y:.4f} Hz<extra></extra>"
        ),
        row=1, col=1
    )
    
    # Reference line at 50Hz
    fig.add_hline(y=50.0, line_dash="dash", line_color="white", opacity=0.3, row=1, col=1)

    # Bottom Plot: VPM
    fig.add_trace(
        go.Scatter(
            x=df_sorted["PRECISE_TIME"],
            y=df_sorted["VPM"],
            name="VPM",
            line=dict(color="#FF4B4B", width=1.2),
            hovertemplate="Time: %{x}<br>VPM: %{y:.2f}<extra></extra>"
        ),
        row=2, col=1
    )

    # Layout Customization
    fig.update_layout(
        height=850,
        template="plotly_dark",
        hovermode="x unified",
        showlegend=False,
        margin=dict(l=50, r=50, t=100, b=50)
    )
    
    # Update y-axis titles
    fig.update_yaxes(title_text="Frequency (HZ)", row=1, col=1)
    fig.update_yaxes(title_text="VPM", row=2, col=1)
    fig.update_xaxes(title_text="Time", row=2, col=1)
    
    # Add range selector buttons and range slider on bottom x-axis
    fig.update_xaxes(
        rangeselector=dict(
            buttons=list([
                dict(count=1, label="1 min", step="minute", stepmode="backward"),
                dict(count=5, label="5 min", step="minute", stepmode="backward"),
                dict(step="all", label="All")
            ])
        ),
        rangeslider=dict(visible=True),
        row=2,
        col=1
    )
    
    return fig

def main():
    st.set_page_config(page_title="Bhadla Oscillation Analysis", layout="wide")
    
    st.title("⚡ Bhadla High-Res Oscillation Analysis")
    st.markdown("This tool processes 20Hz sample data by interpolating sub-second timestamps for clear visualization.")
    
    with st.spinner("Processing data files..."):
        df = load_all_data()

    if df.empty:
        st.error("No data files found in the current directory.")
        return

    # Sidebar Filter & Export
    st.sidebar.header("Settings")
    all_files = sorted(df["SOURCE_FILE"].unique())
    selected_file = st.sidebar.selectbox("Select Data Source", all_files)
    
    display_df = df[df["SOURCE_FILE"] == selected_file]
    
    # Generate the plot
    fig = create_oscillation_plot(display_df)
    
    # 📥 DOWNLOAD SECTION in Sidebar
    st.sidebar.markdown("---")
    st.sidebar.subheader("Export Options")
    
    # Convert figure to HTML string
    html_content = fig.to_html(
        include_plotlyjs='cdn', 
        full_html=True, 
        config={'displaylogo': False}
    )
    
    st.sidebar.download_button(
        label="Download Plot as HTML",
        data=html_content,
        file_name=f"Bhadla_Analysis_{selected_file.split('.')[0]}.html",
        mime="text/html",
        help="Download an interactive version of this chart that opens in any browser."
    )
    
    # Main Dashboard Area
    st.plotly_chart(fig, use_container_width=True)
    
    # Statistics Summary
    st.markdown("### Signal Summary")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Max Freq", f"{display_df['HZ'].max():.3f} Hz")
    c2.metric("Min Freq", f"{display_df['HZ'].min():.3f} Hz")
    c3.metric("Peak-to-Peak", f"{display_df['HZ'].max() - display_df['HZ'].min():.3f} Hz")
    c4.metric("Total Samples", f"{len(display_df):,}")

if __name__ == "__main__":
    main()