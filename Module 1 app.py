import streamlit as st
import pandas as pd
import io

def main():
    st.title("Exam Data Processor")
    st.write("Upload your original Excel file to process the syllabus items and marks.")

    # 1. FILE INPUT (Streamlit File Uploader)
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            with st.spinner("Processing data..."):
                # 2. Read original file
                df = pd.read_excel(uploaded_file)

                # Ensure the dataframe has at least 6 columns before renaming
                if len(df.columns) < 6:
                    st.error("The uploaded file does not have enough columns (minimum 6 required).")
                    return

                # 3. Update headings (Renaming first 6 columns)
                new_headers = {
                    df.columns[0]: 'Qpaper',
                    df.columns[1]: 'QNumber1',
                    df.columns[2]: 'QNumber2',
                    df.columns[3]: 'Level 1',
                    df.columns[4]: 'Level 2',
                    df.columns[5]: 'Marks'
                }
                df.rename(columns=new_headers, inplace=True)

                # 4. Clean Levels and add leading zeros
                for col in ['Level 1', 'Level 2']:
                    if col in df.columns:
                        # Convert to string and strip potential '.0'
                        df[col] = df[col].astype(str).str.replace(r'\.0$', '', regex=True)
                        # Pad with leading zero
                        df[col] = df[col].str.zfill(2)

                # 5. Replicate Access Query 1: Concatenate Levels
                df['Syllabus item'] = df['Level 1'] + "." + df['Level 2']

                # 6. Replicate Access Query 2: Grouping and Aggregation
                final_df = df.groupby(['Qpaper', 'QNumber1']).agg({
                    'Syllabus item': 'max',
                    'Marks': 'sum'
                }).reset_index()

                # 7. Final Formatting (Renaming and Sorting)
                final_df.rename(columns={
                    'Syllabus item': 'MaxOfSyllabus item',
                    'Marks': 'SumOfMarks'
                }, inplace=True)

                final_df.sort_values(by=['Qpaper', 'QNumber1'], inplace=True)

            st.success("Success! Process complete.")
            
            # Preview first 5 rows
            st.write("### Data Preview")
            st.dataframe(final_df.head())

            # 8. Save final result to memory for download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False)
            processed_data = output.getvalue()

            # Provide download button
            st.download_button(
                label="📥 Download Processed Excel File",
                data=processed_data,
                file_name="final_query_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()