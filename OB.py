import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
import io
from io import BytesIO
import openpyxl




#set up data cleaning App
st.set_page_config(page_title="Data Clean APP", layout="wide")
st.title("Data Clean App")
st.write("Cleans data and transform from CSV to Excel formats")

uploaded_files=st.file_uploader("upload your files(CSV or Excel):", type=["csv","xlsx"],accept_multiple_files=True)
if uploaded_files:
    for file in uploaded_files:
        file_ext=os.path.splitext(file.name)[-1].lower()

        if file_ext==".csv":
            df=pd.read_csv(file)

        elif file_ext in ['.xlsx','.xls']:
            df=pd.read_excel(file,engine="openpyxl")
        else: 
            st.error("Unsupported file type:{file_ext} or contact Aubry Madeira on +265881980797")
            continue
        #Display info about the file
        st.write(f"File Name: {file.name}")
        st.write(f"File Size:{file.size/1024}")
        #Show 5 rows of our df
        st.write ("Preview the head of the Dataframe")
        st.dataframe(df.head())

        # Data Inspection
        st.subheader("Data Inspection")
        if st.checkbox(f"Data inspection for {file.name}"):
            col1, col2,col3,col4=st.columns(4)
            with col1:
                if st.button(f"Total Missing Value for {file.name}"):
                    missing_counts = df.isna().sum()
                    st.write("Total missing values: ", missing_counts[missing_counts > 0])

            with col2:
                if st.button(f"Show rows contain missing value for {file.name}"):
                    missing_rows = df[df.isna().any(axis=1)]
                    st.write("\nRows with missing values:")
                    st.dataframe(missing_rows)

            with col3:
                if st.button(f"Total duplicate for {file.name}"):
                    duplicate_rows = df[df.duplicated(keep=False)]
                    st.write("\nNumber of rows involved in duplicates:", len(duplicate_rows))

            with col4:
                if st.button(f"Show duplicated rows for {file.name}"):
                    duplicate_rows = df[df.duplicated(keep=False)]
                    st.write("Duplicate rows:")
                    st.dataframe(duplicate_rows)


        # Options for data Cleaning
        st.subheader("Data Cleaning Options")
        if st.checkbox(f" Do you want to Clean Data for {file.name}"):
            col1,col2=st.columns(2)
            with col1:
                if st.button(f"Remove Duplicates from {file.name}"):
                    df.drop_duplicates(inplace=True)
                    st.write("Duplicates Removed!")
            with col2:
                if st.button(f"Fill Missing Values for {file.name}"):
                    numeric_cols=df.select_dtypes(include=['number']).columns
                    df[numeric_cols] =df[numeric_cols].fillna(df[numeric_cols].mean())
                    st.write("Missing values have been filled")

        # Choose Specific Columns to keep or Convert
        st.subheader("Select Colums to Convert")
        columns=st.multiselect(f"Choose Columns for {file.name}",df.columns,default=df.columns)
        df=df[columns]

        # Create Visualization
        st.subheader("Data Visualization")
        #select Chart type
        if st.checkbox(f"Show Visualization for {file.name}"):
            chart_type=st.selectbox("Select Chart Type",["Bar Chart","Line Chart","Boxplot","Pie Chart","Histogram"])
            # get columns
            all_columns=df.columns.tolist()
            # get numeric columns
            numeric_columns=df.select_dtypes(include='number').columns.tolist()
            # Chart title and axis labels
            st.write("### Chart Labelss")
            chart_title=st.text_input("Enter Chart Title",value="My Chart")
            col1, col2=st.columns(2)
            with col1:
                x_label=st.text_input("X-axis Label",value="X-axis")
            with col2:
                y_label=st.text_input("Y-axis Label",value="Y-axis")
            # BAR Chart
            if chart_type=="Bar Chart":
                st.write("### Bar Chart Settings")
                col1,col2=st.columns(2)
                with col1:
                    x_column=st.selectbox("Select X-axis Column",all_columns)
                with col2:
                    if numeric_columns:
                        y_column=st.selectbox("Select Y-axis Column",numeric_columns)
                    else:
                        st.error("no numeric columns found!")
                        y_column=None

        
                if st.button(f"Plot Bar Chart for {file.name}"):
                    if y_column is not None:
                        fig, ax=plt.subplots()
                        ax.bar(df[x_column].astype(str),df[y_column])
                        ax.set_title(chart_title)
                        ax.set_xlabel(x_label)
                        ax.set_ylabel(y_label)
                        plt.xticks(rotation=45)
                        plt.tight_layout()
                        st.pyplot(fig)
                        st.success("Bar Chart ploted succefully !")

                # Line Chart
            elif chart_type== "Line Chart":
                st.write("### Line Chart Settings")
                col1,col2=st.columns(2)
                with col1:
                        x_column=st.selectbox("Select X-axis Column", all_columns) 
                with col2:
                    if numeric_columns:
                            y_column=st.selectbox("Select Y-axixs Column",numeric_columns)
                    else:
                        st.error("No numeric column found")
                        y_column=None
                                          
                if st.button(f"Plot Linear graph for {file.name}"):
                    if y_column is not None:
                            fig, ax=plt.subplots()
                            ax.plot(df[x_column],df[y_column])
                            ax.set_title(chart_title)
                            ax.set_xlabel(x_label)
                            ax.set_ylabel(y_label)
                            plt.xticks(rotation=45)
                            plt.tight_layout()
                            st.pyplot(fig)
                            st.success("Linear Graph ploted succefully")
                    #Boxplot
            elif chart_type=="Boxplot":
                st.write("### Boxplot Settings")
                if numeric_columns:
                    selected_columns=st.multiselect("Select Numeric Columns", numeric_columns)
                    if st.button(f"Plot Boxplot for {file.name}"):
                        if len(selected_columns)>0:
                            fig,ax=plt.subplots()
                            df[selected_columns].boxplot(ax=ax)
                            ax.set_title(chart_title)
                            ax.set_xlabel(x_label)
                            ax.set_ylabel(y_label)
                            plt.tight_layout()
                            st.pyplot(fig)
                            st.success("Boxplot plotted successfully!")
                        else:
                            st.warning("Please select at least one numeric column")
                    else:
                        st.error("No numeric columns in your Excel file.") 
            # Pie chart
            elif chart_type=="Pie Chart":
                st.write ("### Pie Chart Settings")
                col1, col2=st.columns(2)
                with col1:
                    category_column=st.selectbox("Select Category Column", all_columns)   
                with col2:
                    if numeric_columns:
                        value_column=st.selectbox("Select Value Column", numeric_columns)
                    else:
                        st.error("No Numeric Columns Found")
                        value_column=None
                if st.button("Plot Pie Chart"):
                    if value_column is not None:
                        # Group values by category
                        pie_data=df.groupby(category_column)[value_column].sum()
                        fig, ax=plt.subplots()
                        ax.pie(pie_data.values,labels=pie_data.index,autopct="%1.1f%%",startangle=90)
                        ax.set_title(chart_title)
                        plt.tight_layout()
                        st.pyplot(fig)
                        st.success("Pie chart plotted successfully!")
            # Histogram 
            elif chart_type=="Histogram":
                st.write("### Histogram Settings")
                if numeric_columns:
                    column=st.selectbox("Select Numeric Column", numeric_columns)
                    bins=st.slider("Number of Bins", min_value=5, max_value=100,value=20)
                if st.button("Plot Histogram"):
                    fig, ax=plt.subplots()
                    ax.hist(df[column].dropna(),bins=bins)
                    ax.set_title(chart_title)
                    ax.set_xlabel(x_label)
                    ax.set_ylabel(y_label)
                    plt.tight_layout()
                    st.pyplot(fig)
                    st.success("Histogram Plotted Successfully!")
                else:
                    st.error("No numeric columns found.") 
            # Download Cleaned Excel
            st.subheader("Download Cleaned Excel")
             
            if st.button("Prepare Cleaned Excel"):
                output=BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl")as writer:
                    df.to_excel(writer, index=False,sheet_name="Cleaned Data")
                    output.seek(0)
                st.download_button(label="Download Cleaned Excel", data=output.getvalue(),file_name="cleaned_data.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")                                                 

                        
        