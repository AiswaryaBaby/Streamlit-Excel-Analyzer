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

    # Button to show lowest 5 scores
    if st.button('Lowest 5 Scores'):
        lowest_scores = filtered_data.mean(axis=0).sort_values().head(5)
        st.write("Lowest 5 Scores:")
        st.write(lowest_scores)

    # Button to show highest 5 scores
    if st.button('Highest 5 Scores'):
        highest_scores = filtered_data.mean(axis=0).sort_values(ascending=False).head(5)
        st.write("Highest 5 Scores:")
        st.write(highest_scores)

    # Button for generating Bar Chart
    if st.button('Generate Bar Chart'):
        avg_scores = filtered_data.mean(axis=0)
        plt.figure(figsize=(10,6))
        sns.barplot(x=avg_scores.index, y=avg_scores.values)
        plt.xticks(rotation=90)
        plt.title('Average Scores for Each Statement')
        st.pyplot(plt)

    # Button for generating Pie Chart
    if st.button('Generate Pie Chart'):
        avg_scores = filtered_data.mean(axis=0)
        pie_data = avg_scores.value_counts()
        plt.figure(figsize=(8,8))
        plt.pie(pie_data, labels=pie_data.index, autopct='%1.1f%%')
        plt.title('Distribution of Scores')
        st.pyplot(plt)

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
        for statement in filtered_data.columns:
            avg_score = filtered_data[statement].mean()
            action = recommend_action(avg_score, action_plan_data)
            recommendations.append([statement, avg_score, action])

        # Display recommended actions
        recommendations_df = pd.DataFrame(recommendations, columns=['Statement', 'Average Score', 'Recommended Action'])
        st.write(recommendations_df)

        # Option to download recommendations as Excel
        recommendations_df.to_excel('employee_engagement_recommendations.xlsx', index=False)
        st.download_button(label='Download Recommendations', data=open('employee_engagement_recommendations.xlsx', 'rb'), file_name='employee_engagement_recommendations.xlsx', mime='application/vnd.ms-excel')
