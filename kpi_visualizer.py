import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

def load_data(excel_file):
    """
    Load data from an Excel file into pandas DataFrames.
    
    Parameters:
    - excel_file (str): Path to the Excel file.
    
    Returns:
    - Tuple of pandas DataFrames for each sheet.
    """
    # Load data from Excel file into DataFrames
    df_value_targets = pd.read_excel(excel_file, sheet_name='Tasks')
    df_learning = pd.read_excel(excel_file, sheet_name='learning')
    df_self_improvement = pd.read_excel(excel_file, sheet_name='self_improvement')
    df_public_posts = pd.read_excel(excel_file, sheet_name='public_posts')
    df_enjoyment = pd.read_excel(excel_file, sheet_name='enjoyment')
    
    return df_value_targets, df_learning, df_self_improvement, df_public_posts, df_enjoyment

def create_dashboard(df_value_targets, df_learning, df_self_improvement, df_public_posts, df_enjoyment):
    """
    Create a Plotly dashboard with multiple subplots.

    Parameters:
    - df_value_targets (pd.DataFrame): DataFrame for value vs target data.
    - df_learning (pd.DataFrame): DataFrame for Learning data.
    - df_self_improvement (pd.DataFrame): DataFrame for Self-Improvement data.
    - df_public_posts (pd.DataFrame): DataFrame for Public Posts data.
    - df_enjoyment (pd.DataFrame): DataFrame for Enjoyment data.
    
    Returns:
    - fig (plotly.graph_objs.Figure): Plotly figure object for the dashboard.
    """
    # Define colors for legends
    legend_colors = {
        'Learning Hours': '#1f77b4',  # Blue
        'Self-Improvement Activities': '#ff7f0e',  # Orange
        'Public Posts by Month': '#2ca02c',  # Green
    }
    
    # Create subplots layout
    fig = make_subplots(
        rows=5, cols=2,
        subplot_titles=("Value vs Target - Task 1", "Value vs Target - Task 2", 
                        "Value vs Target - Task 3", "Learning Hours", 
                        "Self-Improvement Activities", "Public Posts by Month", 
                        "Enjoyment Activities", None),
        specs=[[{"type": "pie"}, {"type": "pie"}],
               [{"type": "pie"}, {"type": "bar"}],
               [{"type": "bar"}, {"type": "bar"}],
               [{"type": "bar"}, {"type": "pie"}],
               [{"colspan": 2}, None]]
    )
    
    # Add Pie Charts for Value vs Target for each task
    tasks = df_value_targets['Task']
    values = df_value_targets['value']
    targets = df_value_targets['target']
    
    for i, (task, value, target) in enumerate(zip(tasks, values, targets)):
        fig.add_trace(go.Pie(labels=['Value', 'Target'], 
                             values=[value, target - value],
                             name=f'{task} - Value vs Target'), 
                      row=(i//2)+1, col=(i%2)+1)
    
    # Add Bar Chart with KDE for Learning
    fig.add_trace(go.Bar(x=df_learning['Course'], y=df_learning['Hours'], name='Learning Hours', marker_color=legend_colors['Learning Hours']), 
                  row=2, col=2)
    
    # Add Bar Chart for Self-Improvement
    fig.add_trace(go.Bar(x=df_self_improvement['Activity'], y=df_self_improvement['Hours'], name='Self-Improvement Activities', marker_color=legend_colors['Self-Improvement Activities']), 
                  row=3, col=1)
    
    # Add Bar Chart for Public Posts
    fig.add_trace(go.Bar(x=df_public_posts['Month_Year'], y=df_public_posts['Posts'], name='Public Posts by Month', marker_color=legend_colors['Public Posts by Month']), 
                  row=3, col=2)
    
    # Add Pie Chart for Enjoyment Activities
    fig.add_trace(go.Bar(x=df_enjoyment['Event'], y=df_enjoyment['Count'], name='Enjoyment Activities'), 
                  row=4, col=1)
    
    # Update layout for scrollable and non-overlapping design
    fig.update_layout(
        height=1000,  # Adjust height as needed for scrollability
        showlegend=True,  # Show legend with colors
        legend=dict(
            orientation="h",  # Horizontal legend
            yanchor="bottom",
            y=-0.3,
            xanchor="left",
            x=0
        ),
        title_text="Dashboard with Multiple Plots",
        margin=dict(t=50, b=50),
        hovermode='closest'
    )
    
    return fig

def main():
    """
    Main function to run the dashboard creation process.
    """
    excel_file = 'kpi_data_dashboard.xlsx'
    
    # Load data from Excel file
    df_value_targets, df_learning, df_self_improvement, df_public_posts, df_enjoyment = load_data(excel_file)
    
    # Create dashboard figure
    fig = create_dashboard(df_value_targets, df_learning, df_self_improvement, df_public_posts, df_enjoyment)
    
    # Display the dashboard
    fig.show()

if __name__ == "__main__":
    main()
