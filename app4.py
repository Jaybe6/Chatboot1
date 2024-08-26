import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns


# Function to load data
def load_data(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except FileNotFoundError:
        st.error(f"The file was not found at {file_path}")
        return None
    except Exception as e:
        st.error(f"An error occurred: {e}")
        return None


# Function to display company information
def display_company_info(df, company_name):
    company_data = df[df['Company'] == company_name]
    if company_data.empty:
        st.write(f"<span style='color: white;'>No data found for company: {company_name}</span>",
                 unsafe_allow_html=True)
        return None
    else:
        st.write(f"<span style='color: white;'>Data for company: {company_name}</span>", unsafe_allow_html=True)
        st.write(company_data)
        return company_data


# Function to plot charts
def plot_charts(df, company_name):
    # Filter data for specific company
    company_data = df[df['Company'] == company_name]

    if company_data.empty:
        st.write("<span style='color: white;'>No data to display for the specified company.</span>",
                 unsafe_allow_html=True)
        return

    # Remove duplicate entries for charting
    unique_df = df.drop_duplicates(subset=['Company'])

    # Set a black background for the plots
    plt.style.use('dark_background')

    # Bar Chart: User Input vs. Top 5 Companies by Revenue
    st.subheader('Revenue Model')
    top_5 = unique_df.nlargest(5, 'Revenue')
    top_5_and_user = pd.concat([top_5, company_data]).drop_duplicates(subset=['Company'])

    plt.figure(figsize=(15,8))  # Increase the size of the bar chart
    bar_plot = sns.barplot(x='Company', y='Revenue', data=top_5_and_user, palette='viridis')
    plt.title('Revenue Model', color='white', fontsize=20)
    plt.xticks(rotation=25, ha='right', color='white')  # Ensure column names are horizontal and visible
    plt.tight_layout()

    # Decrease the width of the columns and increase the spacing
    for index, patch in enumerate(bar_plot.patches):
        patch.set_width(0.6)  # Adjust width of bars
        patch.set_edgecolor('black')  # Add black edge color to bars

    # Add data labels to the bar chart with increased width and bold font
    for p in bar_plot.patches:
        bar_plot.annotate(
            format(p.get_height(), '.0f'),
            (p.get_x() + p.get_width() / 2., p.get_height()),
            ha='center', va='center',
            xytext=(0, 10),
            textcoords='offset points',
            color='white',
            fontsize=16,
            fontweight='bold'
        )

    st.pyplot(plt)

    # Pie Chart: User Input vs. Top 5 Companies by Profit
    st.subheader('Order Share')
    top_5_profit = unique_df.nlargest(5, 'Profit')
    top_5_profit_and_user = pd.concat([top_5_profit, company_data]).drop_duplicates(subset=['Company'])

    plt.figure(figsize=(10, 10))  # Increase the size of the pie chart
    plt.pie(top_5_profit_and_user['Profit'], labels=top_5_profit_and_user['Company'], autopct='%1.1f%%',
            colors=sns.color_palette("dark:#5A9_r", len(top_5_profit_and_user)))
    plt.title('Order Share', color='white', fontsize=20)
    plt.tight_layout()
    st.pyplot(plt)

    # Histogram: Distribution of Profit
    st.subheader('Profit Distribution')
    plt.figure(figsize=(10,10))  # Increase the size of the histogram
    bins = 20
    counts, bin_edges, patches = plt.hist(df['Profit'], bins=bins, color='skyblue', edgecolor='black', alpha=0.7,
                                          label='Profit Distribution')
    plt.title('Profit Distribution', color='white', fontsize=20)
    plt.xlabel('Profit', color='white')
    plt.ylabel('Frequency', color='white')

    # Add data labels to the histogram
    for count, bin_edge, patch in zip(counts, bin_edges, patches):
        height = patch.get_height()
        plt.text(
            bin_edge + (patch.get_width() / 2), height,
            f'{int(height)}',
            ha='center', va='bottom',
            color='white',
            fontsize=12,
            fontweight='bold'
        )

    plt.tight_layout()
    st.pyplot(plt)


# Streamlit app
def main():
    st.title('Excel Data Insights Chatbot')

    # Set background color for Streamlit dashboard
    st.markdown(
        """
        <style>
        .main {
            background-color: black;
        }
        .stButton button {
            background-color: #262730;
            color: white;
        }
        .stTextInput input {
            background-color: #262730;
            color: white;
        }
        .stTextInput label {
            color: white;
        }
        .stSubheader {
            color: #FFD700; /* Gold color for better visibility on black */
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # File path for the Excel file
    file_path = "C:\\Users\\Joydip\\Desktop\\Data2.xlsx"

    # Load data from the specified file path
    df = load_data(file_path)

    if df is not None:
        st.write('Data Loaded Successfully!')

        # Chatbot input
        st.subheader('Ask a question:')
        company_name = st.text_input('Enter company name to get insights:')

        if company_name:
            company_data = display_company_info(df, company_name)
            if company_data is not None:
                # Display revenue and profit information for the selected company
                st.write(f"<span style='color: white;'>Company: {company_name}</span>", unsafe_allow_html=True)
                st.write(f"<span style='color: white;'>Revenue: {company_data['Revenue'].values[0]}</span>",
                         unsafe_allow_html=True)
                st.write(f"<span style='color: white;'>Profit: {company_data['Profit'].values[0]}</span>",
                         unsafe_allow_html=True)

                # Plot charts based on the user input
                plot_charts(df, company_name)


if __name__ == '__main__':
    main()
