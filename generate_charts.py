import pandas as pd
import matplotlib.pyplot as plt

# --- 1. SET UP THE STYLE ---
# Using 'seaborn-v0_8-pastel' gives it a clean, professional, soft look.
plt.style.use('seaborn-v0_8-pastel') 

# Define a consistent color palette for high, medium, and low.
colors = {
    'Critical': '#E74C3C', # Vivid Red
    'High': '#F39C12',     # Vibrant Orange
    'Medium': '#3498DB',   # Calm Blue
    'Low': '#2ECC71'       # Fresh Green
}

# --- 2. GENERATE SAMPLE DATA ---
# (In your main script, this will come from the AI's DataFrame)
# Suppose the AI summarized the modes for all 20 tickets:
data = {
    'Mode': ['High', 'Critical', 'Medium', 'Medium', 'Low', 'High', 'Medium', 'Critical', 'Low', 'Medium'],
    'Technical_Domain': ['API', 'UI', 'Database', 'API', 'General', 'Security', 'Database', 'API', 'UI', 'Network']
}
df = pd.DataFrame(data)

# Calculate counts for our visualizations
mode_counts = df['Mode'].value_counts()
# Reorder specifically so the most urgent always appears on top/first:
mode_order = ['Critical', 'High', 'Medium', 'Low']
mode_counts = mode_counts.reindex([m for m in mode_order if m in mode_counts.index]).fillna(0)

domain_counts = df['Technical_Domain'].value_counts()

# --- 3. PIE CHART GENERATION ---
# Focus: Showing the relative size of urgency modes.
def create_pie_chart(counts, palette):
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(aspect="equal"))

    # Create the chart
    wedges, texts, autotexts = ax.pie(
        counts, 
        autopct='%1.1f%%', 
        textprops=dict(color="w"), 
        colors=[palette.get(m, '#95A5A6') for m in counts.index],
        startangle=90, 
        pctdistance=0.8
    )

    # Styling labels and percentages
    ax.legend(
        wedges, 
        counts.index,
        title="Ticket Mode",
        loc="center left",
        bbox_to_anchor=(1, 0, 0.5, 1)
    )
    plt.setp(texts, size=12, fontweight="bold")
    plt.setp(autotexts, size=11, fontweight="bold")
    ax.set_title("Overall Urgency Breakdown", size=16, fontweight="bold")
    
    # Save a high-resolution PNG for Excel/GIF use
    plt.savefig('urgency_pie_chart.png', dpi=300, bbox_inches='tight')
    plt.show() # Close this window to run the next part.

# --- 4. BAR DIAGRAM GENERATION ---
# Focus: Comparing different technical domains.
def create_bar_chart(counts, chart_title, y_label):
    plt.figure(figsize=(10, 6))
    
    bars = plt.bar(counts.index, counts.values, color='#3498DB') # Single color for clean comparison
    
    # Label styling
    plt.xlabel("Technical Area", fontsize=12, fontweight='bold')
    plt.ylabel(y_label, fontsize=12, fontweight='bold')
    plt.title(chart_title, fontsize=16, fontweight='bold')
    
    # Rotating x-axis labels if needed
    plt.xticks(rotation=45, ha='right')
    
    # Grid lines and saving
    plt.grid(axis='y', linestyle='--', alpha=0.5)
    plt.savefig('domain_bar_chart.png', dpi=300, bbox_inches='tight')
    plt.show()

# --- 5. EXECUTE ---
create_pie_chart(mode_counts, colors)
create_bar_chart(domain_counts, "Ticket Frequency by Domain", "Number of Tickets")

print("Both charts generated successfully as high-resolution PNGs!")