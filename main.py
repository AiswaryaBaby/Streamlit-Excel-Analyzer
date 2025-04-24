import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Title of the app
st.title('Employee Engagement Data Analysis')

# Brief description of the app
st.write("This app allows you to analyze Employee Engagement scores based on various filters and generate insights with action plans.")

# File uploader for the Employee Engagement Data (Excel file)
engagement_file = st.file_uploader("Upload Employee Engagement Data (Excel file)", type=["xlsx"])
if engagement_file is not None:
    # Load the Excel file into pandas dataframe
    engagement_data = pd.read_excel(engagement_file, engine='openpyxl')
    st.write("Employee Engagement Data:")
    st.write(engagement_data.head())  # Show the first few rows of the uploaded data

# File uploader for the Action Plan Data (Excel file)
action_plan_file = st.file_uploader("Upload Action Plan Data (Excel file)", type=["xlsx"])
if action_plan_file is not None:
    # Load the Action Plan Excel file into pandas dataframe
    action_plan_data = pd.read_excel(action_plan_file, engine='openpyxl')
    st.write("Action Plan Data:")
    st.write(action_plan_data.head())  # Show the first few rows of the action plan data

# Only proceed with analysis if both files are uploaded
if engagement_file is not None and action_plan_file is not None:

    # Dropdown for BUHR Name
    buhr_name = st.selectbox('Select BUHR NAME:', engagement_data['BUHR NAME'].unique())

    # Dropdown for Department
    department_name = st.selectbox('Select Department:', engagement_data['Department'].unique())

    # Dropdown for BUHead Name
    buhead_name = st.selectbox('Select BUHEAD NAME:', engagement_data['BUHEAD NAME'].unique())

    # Filter the engagement data based on the selections
    filtered_data = engagement_data[(engagement_data['BUHR NAME'] == buhr_name) &
                                    (engagement_data['Department'] == department_name) &
                                    (engagement_data['BUHEAD NAME'] == buhead_name)]

    st.write(f"Displaying data for {buhr_name} in {department_name} led by {buhead_name}")
    st.write(filtered_data)

    # --- Functionality for Lowest Scoring Statements ---
    if st.checkbox('Show Lowest Scoring Statements Visually (Q1-Q26)'):
        # Identify statement columns (assuming they start with 'Q-' and are within the range)
        statement_columns = [col for col in filtered_data.columns if col.startswith('Q-') and 1 <= int(col.split('-')[1]) <= 26]

        if statement_columns:
            # Calculate the average score for each statement
            avg_statement_scores = filtered_data[statement_columns].mean().sort_values(ascending=True)

            # Display the top N lowest scoring statements (e.g., top 5)
            num_lowest = st.slider("Number of Lowest Statements to Show", min_value=1, max_value=len(avg_statement_scores), value=5)
            lowest_n_scores = avg_statement_scores.head(num_lowest)

            # Create a bar plot for the lowest scoring statements
            plt.figure(figsize=(10, 6))
            sns.barplot(x=lowest_n_scores.index, y=lowest_n_scores.values)
            plt.xlabel('Statement')
            plt.ylabel('Average Score')
            plt.title(f'Top {num_lowest} Lowest Scoring Statements (Q1-Q26)')
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            st.pyplot(plt)
        else:
            st.warning("No statements found within the Q1-Q26 range for the selected data.")

    # --- Existing Functionality ---
    # Button to show lowest 5 scores (overall)
    if st.button('Lowest 5 Scores (Overall)'):
        numeric_cols = filtered_data.select_dtypes(include=['number']).columns
        if not numeric_cols.empty:
            lowest_scores = filtered_data[numeric_cols].mean(axis=0).sort_values().head(5)
            st.write("Lowest 5 Scores (Overall):")
            st.write(lowest_scores)
        else:
            st.warning("No numeric columns found to calculate scores.")

    # Button to show highest 5 scores (overall)
    if st.button('Highest 5 Scores (Overall)'):
        numeric_cols = filtered_data.select_dtypes(include=['number']).columns
        if not numeric_cols.empty:
            highest_scores = filtered_data[numeric_cols].mean(axis=0).sort_values(ascending=False).head(5)
            st.write("Highest 5 Scores (Overall):")
            st.write(highest_scores)
        else:
            st.warning("No numeric columns found to calculate scores.")

    # Button for generating Bar Chart (all statements)
    if st.button('Generate Bar Chart (All Statements)'):
        numeric_cols = filtered_data.select_dtypes(include=['number']).columns
        if not numeric_cols.empty:
            avg_scores = filtered_data[numeric_cols].mean(axis=0)
            plt.figure(figsize=(10, 6))
            sns.barplot(x=avg_scores.index, y=avg_scores.values)
            plt.xticks(rotation=90)
            plt.title('Average Scores for All Statements')
            st.pyplot(plt)
        else:
            st.warning("No numeric columns found to generate the bar chart.")

    # Button for generating Pie Chart (all statements)
    if st.button('Generate Pie Chart (All Statements)'):
        numeric_cols = filtered_data.select_dtypes(include=['number']).columns
        if not numeric_cols.empty:
            avg_scores = filtered_data[numeric_cols].mean(axis=0)
            pie_data = avg_scores.value_counts()
            plt.figure(figsize=(8, 8))
            plt.pie(pie_data, labels=pie_data.index, autopct='%1.1f%%')
            plt.title('Distribution of Scores (All Statements)')
            st.pyplot(plt)
        else:
            st.warning("No numeric columns found to generate the pie chart.")

    # Function to recommend actions based on score threshold
    def recommend_action(score, action_plan_data):
        if score >= 4.0:
            action = action_plan_data[action_plan_data['Score'] == 'High']['Action'].values[0]
        elif score <= 2.0:
            action = action_plan_data[action_plan_data['Score'] == 'Low']['Action'].values[0]
        else:
            action = action_plan_data[action_plan_data['Score'] == 'Medium']['Action'].values[0]
        return action

    # Button to get recommended actions
    if st.button('Generate Recommended Actions'):
        recommendations = []
        statement_columns = [col for col in filtered_data.columns if col.startswith('Q-')]
        for statement in statement_columns:
            if statement in filtered_data.columns and pd.api.types.is_numeric_dtype(filtered_data[statement]):
                avg_score = filtered_data[statement].mean()
                action = recommend_action(avg_score, action_plan_data)
                recommendations.append([statement, avg_score, action])

        if recommendations:
            # Display recommended actions
            recommendations_df = pd.DataFrame(recommendations, columns=['Statement', 'Average Score', 'Recommended Action'])
            st.write(recommendations_df)

            # Option to download recommendations as Excel
            recommendations_df.to_excel('employee_engagement_recommendations.xlsx', index=False)
            with open('employee_engagement_recommendations.xlsx', 'rb') as f:
                st.download_button(label='Download Recommendations', data=f, file_name='employee_engagement_recommendations.xlsx', mime='application/vnd.ms-excel')
        else:
            st.warning("No valid statement scores found to generate recommendations.")
