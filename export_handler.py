import pandas as pd
import io
import datetime
import base64
import matplotlib.pyplot as plt
from io import BytesIO

class ExportHandler:
    def __init__(self, data, engine_type):
        self.data = data
        self.engine_type = engine_type
        self.timestamp = datetime.datetime.now().strftime("%Y-%m-%d")
        self.vessel_name = data['Vessel'].iloc[0] if 'Vessel' in data.columns and not data.empty else "Vessel"

    def plot_to_base64(self, fig):
        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches='tight')
        buf.seek(0)
        encoded = base64.b64encode(buf.read()).decode("utf-8")
        buf.close()
        return f'<img src="data:image/png;base64,{encoded}" style="max-width:100%;">'

    def export_all_tabs_to_html(self, tab_data_dict, totaljobs=0, total_missing_jobs=0,
                                total_machinery=0, missing_machinery=0, vesselname="Vessel",
                                criticaljobscount=0, main_engine_jobs=0, ae_jobs=0,
                                job_status_df=None):
        from report_styler import ReportStyler

        styler = ReportStyler()
        css = styler.generate_html_styles(
            styler.get_color_scheme(),
            styler.get_font_settings(),
            styler.get_table_settings()
        )

        html = f"""
        <html>
        <head>
            <meta charset='UTF-8'>
            <style>
            {css}
            .nav-grid {{
                display: flex;
                flex-wrap: wrap;
                gap: 10px;
                margin: 10px 0 30px 0;
                justify-content: space-between;
            }}
            .nav-card {{
                flex: 1 1 calc(16.666% - 20px);
                background: #f1f1f1;
                padding: 10px;
                text-align: center;
                border-radius: 8px;
                box-shadow: 1px 1px 4px rgba(0, 0, 0, 0.1);
            }}
            .nav-card a {{
                text-decoration: none;
                color: #007bff;
                font-weight: bold;
            }}
            #homeButton {{
                position: fixed; bottom: 30px; right: 30px; z-index: 1000;
                background-color: #007bff; color: white; padding: 10px 15px;
                border: none; border-radius: 5px; text-decoration: none;
                font-size: 14px; box-shadow: 2px 2px 5px rgba(0,0,0,0.3);
            }}
            .metric-summary {{
                display: flex; flex-wrap: wrap; justify-content: space-around;
                margin-bottom: 20px;
            }}
            .metric-card {{
                background: #f8f9fa; padding: 20px; margin: 10px;
                border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                flex: 1 1 200px; text-align: center;
            }}
            .metric-title {{ font-weight: bold; color: #333; }}
            .metric-value {{ font-size: 1.5em; color: #007bff; }}
            </style>
        </head>

        <body>
            <a name="top"></a>
            <h1 style='text-align:center;'>Vessel Maintenance Report</h1>
            <hr>
            <div class="metric-summary">
                <div class="metric-card"><div class="metric-title">🛳️ Vessel</div><div class="metric-value">{vesselname}</div></div>
                <div class="metric-card"><div class="metric-title">🧾 Total Jobs</div><div class="metric-value">{totaljobs}</div></div>
                <div class="metric-card"><div class="metric-title">🚨 Critical Jobs</div><div class="metric-value">{criticaljobscount}</div></div>
                <div class="metric-card"><div class="metric-title">❌ Total Missing Jobs</div><div class="metric-value">{total_missing_jobs}</div></div>
                <div class="metric-card"><div class="metric-title">🔧 Missing Machinery</div><div class="metric-value">{missing_machinery}</div></div>
            </div>
        """

        if job_status_df is not None and not job_status_df.empty:
            job_status_dict = dict(zip(job_status_df['Job Status'], job_status_df['Count']))
            total_status_records = job_status_df['Count'].sum()
            html += f"""
            <div class="metric-summary">
                <div class="metric-card"><div class="metric-title">📊 Job Status Records</div><div class="metric-value">{total_status_records}</div></div>
                <div class="metric-card"><div class="metric-title">🆕 New</div><div class="metric-value">{job_status_dict.get('New', 0)}</div></div>
                <div class="metric-card"><div class="metric-title">🕓 Pending</div><div class="metric-value">{job_status_dict.get('Pending', 0)}</div></div>
                # <div class="metric-card"><div class="metric-title">✅ Done</div><div class="metric-value">{job_status_dict.get('Done', 0)}</div></div>
                # <div class="metric-card"><div class="metric-title">🔎 Verified</div><div class="metric-value">{job_status_dict.get('Verified', 0)}</div></div>
                # <div class="metric-card"><div class="metric-title">❔ Other</div><div class="metric-value">{job_status_dict.get('Other', 0)}</div></div>
            </div>
            """
        html += """
            <h2>Navigation</h2>
            <div class="nav-grid">
        """



        for tab in tab_data_dict.keys():
            anchor = tab.replace(" ", "_").replace("(", "").replace(")", "")
            html += f'<div class="nav-card"><a href="#{anchor}">{tab}</a></div>'
        html += "</div><hr>"

        




        for tab, tables in tab_data_dict.items():
            anchor = tab.replace(" ", "_").replace("(", "").replace(")", "")
            html += f'<h2 id="{anchor}">{tab}</h2>'

            if tab == "QuickView Summary":
                try:
                    missing_jobs_df = tables[0]

                    fig1, ax1 = plt.subplots(figsize=(4, 4))
                    wedges1, texts1, autotexts1 = ax1.pie(
                        [totaljobs, total_missing_jobs],
                        labels=["Total Jobs", "Missing Jobs"],
                        autopct=lambda pct: f'{int(pct * (totaljobs + total_missing_jobs) / 100)} ({pct:.1f}%)'
                    )
                    ax1.axis("equal")
                    ax1.set_title("Total Jobs vs Missing Jobs")
                    html += self.plot_to_base64(fig1)
                    plt.close(fig1)

                    fig2, ax2 = plt.subplots(figsize=(4, 4))
                    wedges2, texts2, autotexts2 = ax2.pie(
                        [total_machinery, missing_machinery],
                        labels=["Present", "Missing"],
                        autopct=lambda pct: f'{int(pct * (total_machinery + missing_machinery) / 100)} ({pct:.1f}%)'
                    )
                    ax2.axis("equal")
                    ax2.set_title("Machinery Summary")
                    html += self.plot_to_base64(fig2)
                    plt.close(fig2)

                    fig3, ax3 = plt.subplots(figsize=(18, 6))
                    bars = ax3.bar(missing_jobs_df["Machinery System"], missing_jobs_df["Missing Jobs Count"], color='#5DADE2')
                    for bar in bars:
                        height = bar.get_height()
                        ax3.text(bar.get_x() + bar.get_width() / 2, height + 0.5, str(int(height)), ha='center', va='bottom', fontsize=9)
                    ax3.set_title("Missing Jobs by Machinery System")
                    ax3.set_xlabel("Machinery System")
                    ax3.set_ylabel("Missing Jobs Count")
                    ax3.tick_params(axis='x', rotation=45, labelsize=9)
                    ax3.grid(axis='y', linestyle='--', alpha=0.5)
                    fig3.tight_layout()
                    html += self.plot_to_base64(fig3)
                    plt.close(fig3)
                except Exception as chart_err:
                    html += f"<p>Chart generation failed: {chart_err}</p>"

            for i, df in enumerate(tables):
                if isinstance(df, pd.DataFrame) and not df.empty:
                    if "Main Engine" in tab:
                        titles = [
                            "Maintenance Data for Main Engine",
                            "Main Engine Cylinder Unit Analysis",
                            "Reference Analysis for Main Engine",
                            "Missing Jobs for Main Engine",
                            "Component Status Analysis for Main Engine",
                            "Number of Missing Components for Main Engine"
                        ]
                        title = titles[i] if i < len(titles) else f"Main Engine Table {i+1}"
                    elif "Auxiliary Engine" in tab:
                        titles = [
                            "Task Count Analysis for Auxiliary Engine",
                            "Component Distribution for Auxiliary Engine",
                            "Component Status Analysis for Auxiliary Engine",
                            "Reference Analysis for Auxiliary Engine",
                            "Missing Jobs for Auxiliary Engine"
                        ]
                        title = titles[i] if i < len(titles) else f"Auxiliary Engine Table {i+1}"
                    elif "Machinery Location Analysis" in tab:
                        titles = [
                                "Missing Machinery on Vessel Analysis",
                                "Different Machinery on Vessel Analysis"
                        ]
                        title = titles[i] if i < len(titles) else f"Machinery Location Analysis{i+1}"

                    elif "Purifier Analysis" in tab:
                        titles = [
                            "Task Count Analysis for Purifiers",
                            "Reference Jobs for Purifiers",
                            "Missing Jobs for Purifiers"
                        ]
                        title = titles[i] if i < len(titles) else f"Purifier Table {i+1}"

                    elif "BWTS Analysis" in tab:
                        titles = [
                            "Task Count Analysis for BWTS",
                            "Reference Jobs for BWTS",
                            "Missing Jobs for BWTS"
                        ]
                        title = titles[i] if i < len(titles) else f"BWTS Table {i+1}"

                    
                    elif "Hatch Analysis" in tab:
                        titles = [
                            "Task Count Analysis for Hatches",
                            "Reference Jobs for Hatches",
                            "Missing Jobs for Hatches"
                        ]
                        title = titles[i] if i < len(titles) else f"Hatch Table {i+1}"
                    
                    elif "Cargo Pumping System Analysis" in tab:    
                        titles = [
                            "Task Count Analysis for Cargo Pumping",
                            "Reference Jobs for Cargo Pumping",
                            "Missing Jobs for Cargo Pumping"
                        ]
                        title = titles[i] if i < len(titles) else f"Cargo Pumping Table {i+1}"

                    elif "Inert Gas System Analysis" in tab:    
                        titles = [
                            "Task Count Analysis for Inert Gas System",
                            "Reference Jobs for Inert Gas System",
                            "Missing Jobs for Inert Gas System"
                        ]
                        title = titles[i] if i < len(titles) else f"Inert Gas System Table {i+1}"

                    elif "Cargo Handling System Analysis" in tab:    
                        titles = [
                            "Task Count Analysis for Cargo Handling System",
                            "Reference Jobs for Cargo Handling System",
                            "Missing Jobs for Cargo Handling System"
                        ]
                        title = titles[i] if i < len(titles) else f"Cargo Handling System Analysis {i+1}"

                    elif "Cargo Venting System Analysis" in tab:    
                        titles = [
                            "Task Count Analysis for Cargo Venting System",
                            "Reference Jobs for Cargo Venting System",
                            "Missing Jobs for Cargo Venting System"
                        ]
                        title = titles[i] if i < len(titles) else f"Cargo Venting System Analysis {i+1}"

                    elif "LSA/FFA System Analysis" in tab:    
                        titles = [
                            "Task Count Analysis for LSA/FFA System",
                            "Reference Jobs for LSA/FFA System",
                            "Missing Jobs for LSA/FFA System Analysis"
                        ]
                        title = titles[i] if i < len(titles) else f"LSA/FFA System Analysis {i+1}"
                    
                    elif "Fire Fighting System Analysis" in tab:    
                        titles = [
                            "Task Count Analysis for Fire Fighting System Analysis",
                            "Reference Jobs for Fire Fighting System Analysis",
                            "Missing Jobs for Fire Fighting System Analysis"
                        ]
                        title = titles[i] if i < len(titles) else f"Fire Fighting System Analysis {i+1}"
                    
                    elif "Pump System Analysis" in tab:    
                        titles = [
                            "Task Count Analysis for Pump System Analysis",
                            "Summary Jobs for Pump System Analysis",
                                                ]
                        title = titles[i] if i < len(titles) else f"Pump System Analysis {i+1}"


                    elif "Compressor System Analysis" in tab:
                        titles = [
                            "Matched Compressor Job Code Summary Table",
                            "Missing Job Codes from Compressor Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Compressor Table {i+1}"

                    elif "Ladder System Analysis" in tab:
                        titles = [
                            "Matched Ladder Job Code Summary Table",
                            "Missing Job Codes from Ladder Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Ladder Table {i+1}"


                    elif "Boat System Analysis" in tab:
                        titles = [
                            "Matched Boat Job Code Summary Table",
                            "Missing Job Codes from Boat Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Boat Table {i+1}"


                    elif "Mooring System Analysis" in tab:
                        titles = [
                            "Matched Mooring Job Code Summary Table",
                            "Missing Job Codes from Mooring Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Mooring Table {i+1}"


                    elif "Steering System Analysis" in tab:
                        titles = [
                            "Matched Steering Job Code Summary Table",
                            "Missing Job Codes from Steering Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Steering Table {i+1}"


                    elif "Incinerator System Analysis" in tab:
                        titles = [
                            "Matched Incinerator Job Code Summary Table",
                            "Missing Job Codes from Incinerator Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Incinerator Table {i+1}"




                    elif "STP System Analysis" in tab:
                        titles = [
                            "Matched STP Job Code Summary Table",
                            "Missing Job Codes from STP Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"STP Table {i+1}"

                    elif "OWS System Analysis" in tab:
                        titles = [
                            "Matched OWS Job Code Summary Table",
                            "Missing Job Codes from OWS Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"OWS Table {i+1}"


                    elif "Power Distribution System Analysis" in tab:
                        titles = [
                            "Matched Power Distribution Job Code Summary Table",
                            "Missing Job Codes from Power Distribution Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Power Distribution Table {i+1}"

                    elif "Crane System Analysis" in tab:
                        titles = [
                            "Matched Crane Job Code Summary Table",
                            "Missing Job Codes from Crane Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Crane Table {i+1}"

                    elif "Emergency Generator System Analysis" in tab:
                        titles = [
                            "Matched Emergency Generator Job Code Summary Table",
                            "Missing Job Codes from Emergency Generator Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Emergency Generator Table {i+1}"

                    elif "Bridge System Analysis" in tab:
                        titles = [
                            "Matched Bridge Job Code Summary Table",
                            "Missing Job Codes from Bridge Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Bridge Table {i+1}"

                    elif "Reefer & AC System Analysis" in tab:
                        titles = [
                            "Matched Reefer & AC Job Code Summary Table",
                            "Missing Job Codes from Reefer & AC Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Reefer & AC Table {i+1}"

                    elif "Fan System Analysis" in tab:
                        titles = [
                            "Matched Fan Job Code Summary Table (By Title)",
                            "Total Fan Jobs by Machinery and Sub Component"
                        ]
                        title = titles[i] if i < len(titles) else f"Fan Table {i+1}"

                    elif "Tank System Analysis" in tab:
                        titles = [
                            "Matched Tank Job Code Summary Table",
                            "Missing Job Codes from Tank Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Tank Table {i+1}"

                    elif "FWG & Hydrophore System Analysis" in tab:
                        titles = [
                            "Matched FWG & Hydrophore Job Code Summary Table",
                            "Missing Job Codes from FWG Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"FWG Table {i+1}"

                    elif "Workshop System Analysis" in tab:
                        titles = [
                            "Matched Workshop Job Code Summary Table",
                            "Missing Job Codes from Workshop Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Workshop Table {i+1}"

                    elif "Boiler System Analysis" in tab:
                        titles = [
                            "Matched Boiler Job Code Summary Table",
                            "Missing Job Codes from Boiler Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Boiler Table {i+1}"

                    elif "Miscellaneous System Analysis" in tab:
                        titles = [
                            "Matched Misc Job Code Summary by Function",
                            "Total Misc Jobs by Title",
                            "Missing Job Codes from Misc Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Misc Table {i+1}"

                    elif "Battery System Analysis" in tab:
                        titles = [
                            "Matched Battery Job Code Summary Table",
                            "Missing Job Codes from Battery Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Battery Table {i+1}"

                    elif "Bow Thruster (BT) System Analysis" in tab:
                        titles = [
                            "Matched BT Job Code Summary Table",
                            "Missing Job Codes from BT Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"BT Table {i+1}"

                    elif "LPSCR System Analysis" in tab:
                        titles = [
                            "Matched LPSCR Job Code Summary Table",
                            "Missing Job Codes from LPSCR Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"LPSCR Table {i+1}"

                    elif "HPSCR System Analysis" in tab:
                        titles = [
                            "Matched HPSCR Job Code Summary Table",
                            "Missing Job Codes from HPSCR Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"HPSCR Table {i+1}"


                    elif "LSA Mapping Analysis" in tab:
                        titles = [
                            "LSA Mapping Summary by Function",
                            "Total Mapped LSA Jobs by Title",
                            "Missing Job Codes from LSA Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"LSA Mapping Table {i+1}"

                    elif "FFA Mapping Analysis" in tab:
                        titles = [
                            "FFA Mapping Summary by Function",
                            "Total Mapped FFA Jobs by Title",
                            "Missing Job Codes from FFA Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"FFA Mapping Table {i+1}"

                    elif "Inactive Mapping Analysis" in tab:
                        titles = [
                            "Inactive Mapping Summary by Function",
                            "Total Mapped Inactive Jobs by Title",
                            "Missing Job Codes from Inactive Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Inactive Mapping Table {i+1}"

                    elif "Critical Jobs Mapping Analysis" in tab:
                        titles = [
                            "Critical Mapping Summary by Function",
                            "Total Mapped Critical Jobs by Title",
                            "Missing Job Codes from Critical Reference"
                        ]
                        title = titles[i] if i < len(titles) else f"Critical Mapping Table {i+1}"








                    else:
                        title = f"Table {i+1}"

                    try:
                        if isinstance(df, pd.DataFrame):
                            def color_cells(val):
                                try:
                                    val = float(val)
                                    if val == 0:
                                        return "background-color: #dc3545"
                                    elif val == 1:
                                        return "background-color: #28a745"
                                    elif val > 5:
                                        return "background-color: #fd7e14"
                                except:
                                    return ""

                            numeric_cols = df.select_dtypes(include='number').columns
                            styled_df = df.style.map(color_cells, subset=numeric_cols)
                            html += f"<h4>{title}</h4>" + styled_df.to_html(index=False, border=0)
                        else:
                            html += f"<h4>{title}</h4><p><code>{df}</code></p>"
                    except Exception:
                        html += f"<h4>{title}</h4><p>⚠️ Could not render this section.</p>"


            html += '<a href="#top" id="homeButton">Home</a><hr>'

        html += "</body></html>"
        return html
