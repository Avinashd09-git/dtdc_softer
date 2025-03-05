from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from sqlalchemy import create_engine, text
import os
import pandas as pd
from werkzeug.utils import secure_filename
import shutil 


app = Flask(__name__)
app.secret_key = 'supersecretkey'  # For flashing messages
DATABASE_URL = "postgresql://postgres:hr@localhost:5432/postgres"  # Update with your credentials
engine = create_engine(DATABASE_URL)

UPLOAD_FOLDER = 'uploads'
MERGE_FOLDER = 'merged'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(MERGE_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/operation1', methods=['GET', 'POST'])
def operation1():
    return render_template('Ts.html')

@app.route('/operation2', methods=['GET', 'POST'])
def operation2():
    if request.method == 'POST':
        if 'folder' not in request.files:
            flash("No folder uploaded!", "danger")
            return redirect(request.url)

        files = request.files.getlist('folder')
        if not files:
            flash("No files selected!", "danger")
            return redirect(request.url)

        # Get the original folder name from the first file
        original_folder_name = files[0].filename.split('/')[0]
        folder_name = secure_filename(original_folder_name)
        upload_subfolder = os.path.join(MERGE_FOLDER, folder_name)
        
        if os.path.exists(upload_subfolder):
            shutil.rmtree(upload_subfolder)
        os.makedirs(upload_subfolder)

        for file in files:
            filename = secure_filename(file.filename)
            file_path = os.path.join(upload_subfolder, filename)
            file.save(file_path)

        merged_file_path, table_name = process_excel_folder(upload_subfolder)

        if merged_file_path:
            flash(f"✅ Successfully processed and saved as '{merged_file_path}'.", "success")
            return send_file(merged_file_path, as_attachment=True)
        else:
            flash("⚠️ No valid Excel files found for processing.", "warning")

    return render_template('Software Ts.html')

def process_excel_folder(folder_path):
    """Processes the uploaded Excel files, merges them, and inserts them into PostgreSQL."""
    # Get the original folder name from the path
    folder_name = os.path.basename(os.path.normpath(folder_path))
    
    # Create output file path with folder name
    output_file = os.path.join(MERGE_FOLDER, f"{folder_name}.xlsx")
    all_data = []

    # Loop through all files in the uploaded folder
    for file in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file)

        try:
            # Open file in binary mode to check format
            with open(file_path, "rb") as f:
                header = f.read(8)

            # Skip invalid files (HTML, ZIP, etc.)
            if header.startswith(b'<') or header.startswith(b'PK'):
                try:
                    df = pd.read_html(file_path)[0]
                except Exception as e:
                    print(f"❌ Skipping {file}, cannot read as HTML: {e}")
                    continue
            else:
                df = pd.read_excel(file_path)

            all_data.append(df)

        except Exception as e:
            print(f"❌ Error reading {file}: {e}")

    # Merge and save data if valid files exist
    if all_data:
        merged_df = pd.concat(all_data, ignore_index=True)
        merged_df.to_excel(output_file, sheet_name=folder_name, index=False)

        try:
            # Use folder name for the table name
            merged_df.to_sql(folder_name, engine, if_exists="replace", index=False)
            print(f"✅ Data inserted into PostgreSQL table '{folder_name}'")
        except Exception as e:
            print(f"❌ Failed to insert into PostgreSQL: {e}")

        return output_file, folder_name

    return None, None



@app.route('/operation3', methods=['GET', 'POST'])
def operation3():
    if request.method == 'POST':
        if 'folder' not in request.files:
            flash('No files selected', 'danger')
            return redirect(request.url)

        files = request.files.getlist('folder')

        if not files or files[0].filename == '':
            flash('No files selected', 'danger')
            return redirect(request.url)

        # Automatically derive table name from the first file's folder name
        folder_name = files[0].filename.split('/')[0]
        folder_name = secure_filename(folder_name)

        if not folder_name:
            flash('Could not determine folder name for database table', 'danger')
            return redirect(request.url)

        # Create a temporary folder for the uploaded files
        folder_path = os.path.join(UPLOAD_FOLDER, folder_name)
        os.makedirs(folder_path, exist_ok=True)

        # Save all uploaded files
        excel_count = 0
        for file in files:
            if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
                filename = secure_filename(file.filename)
                file.save(os.path.join(folder_path, filename))
                excel_count += 1

        if excel_count == 0:
            flash('No Excel files were found in the selection', 'danger')
            return redirect(request.url)

        try:
            # Process the files
            output_file = merge_international_ts_files(folder_path, folder_name)
            flash(f'Successfully merged {excel_count} files and uploaded to database table: {folder_name}', 'success')

            # Offer the merged file for download
            return send_file(output_file, as_attachment=True, download_name=f"{folder_name}.xlsx")
        except Exception as e:
            flash(f'Error processing files: {str(e)}', 'danger')
            return redirect(request.url)

    return render_template('International TS.html')



def merge_international_ts_files(folder_path, table_name):
    """Merge Excel files from the International TS folder and upload to PostgreSQL"""
    # Define output file path
    output_file = os.path.join(MERGE_FOLDER, f"{table_name}.xlsx")
    
    # Create an empty DataFrame to store the merged data
    merged_df = pd.DataFrame()
    
    # Loop through all Excel files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
            file_path = os.path.join(folder_path, file_name)
            try:
                # Read third sheet (index 2)
                df = pd.read_excel(file_path, sheet_name=2)
                # Remove last row
                df = df.iloc[:-1]
                # Concatenate to the merged DataFrame
                merged_df = pd.concat([merged_df, df], ignore_index=True)
            except Exception as e:
                app.logger.error(f"Error processing file {file_name}: {str(e)}")
    
    if not merged_df.empty:
        # Remove "Sl.No." column if it exists
        if "Sl.No." in merged_df.columns:
            merged_df.drop(columns=["Sl.No."], inplace=True)
        
        # Remove duplicate rows
        merged_df = merged_df.drop_duplicates().reset_index(drop=True)
        
        # Save merged data to an Excel file
        merged_df.to_excel(output_file, index=False)
        
        # Insert into PostgreSQL
        merged_df.to_sql(table_name, engine, if_exists='replace', index=False)
        
        return output_file
    else:
        raise Exception("No data found to merge.")


@app.route('/operation4', methods=['GET', 'POST'])
def operation4():
    if request.method == 'POST':
        files = request.files.getlist("folder")
        
        if not files:
            flash("❌ No files uploaded.", "danger")
            return redirect(url_for('operation4'))
        
        # Extract common folder name from the first uploaded file
        first_filename = files[0].filename
        folder_name = first_filename.split('/')[0] if '/' in first_filename else "merged_file"

        folder_path = os.path.join(UPLOAD_FOLDER, folder_name)
        os.makedirs(folder_path, exist_ok=True)

        for file in files:
            filename = secure_filename(file.filename.split('/')[-1])  # Get the actual file name
            file.save(os.path.join(folder_path, filename))

        output_file = os.path.join(MERGE_FOLDER, f"{folder_name}.xlsx")  # Use extracted folder name

        # Find all Excel files in the uploaded folder
        excel_files = [f for f in os.listdir(folder_path) if f.endswith(('.xls', '.xlsx'))]
        dfs = []
        
        for file in excel_files:
            file_path = os.path.join(folder_path, file)
            try:
                with open(file_path, "rb") as f:
                    content = f.read(1024)  # Read first 1024 bytes
                
                if b'<html' in content.lower() or b'<table' in content.lower():
                    df = pd.read_html(file_path)[0]  # Process as an HTML file
                else:
                    if file.endswith(".xls"):
                        df = pd.read_excel(file_path, engine="xlrd")  # Process as .xls
                    else:
                        df = pd.read_excel(file_path, engine="openpyxl")  # Process as .xlsx
                
                if "Source_File" in df.columns:
                    df = df.drop(columns=["Source_File"])  # Remove "Source_File" if it exists
                
                dfs.append(df)  # ✅ Append dataframe to list
            
            except Exception as e:
                print(f"❌ Error reading {file}: {e}")
        
        if dfs:
            merged_df = pd.concat(dfs, ignore_index=True)
            # Remove duplicate rows
            merged_df = merged_df.drop_duplicates().reset_index(drop=True)
            merged_df.to_excel(output_file, index=False, engine="openpyxl")
            flash(f"✅ Successfully merged {len(dfs)} files into {folder_name}.xlsx", "success")
            return send_file(output_file, as_attachment=True)
        else:
            flash("❌ No valid Excel files found to merge.", "danger")
            return redirect(url_for('operation4'))
    
    return render_template('Booking.html')


@app.route('/upload', methods=['POST'])
def upload_folder():
    if 'folder' not in request.files:
        flash('No folder selected!', 'danger')
        return redirect(url_for('index')) 

    # Get list of files from the upload
    files = request.files.getlist('folder')
    folder_name = secure_filename(request.files['folder'].filename.split('/')[0])  # Extract the folder name
    folder_path = os.path.join(UPLOAD_FOLDER, folder_name)
    os.makedirs(folder_path, exist_ok=True)

    for file in files:
        file.save(os.path.join(folder_path, secure_filename(file.filename)))

    # Call the merge function
    output_file = merge_excel_files(folder_path, folder_name)

    if output_file:
        # Extract the base file name (without the extension)
        base_file_name = os.path.splitext(os.path.basename(output_file))[0]

        # Load the merged Excel file into a DataFrame
        merged_df = pd.read_excel(output_file)

        # Call the insert_into_postgresql function
        insert_into_postgresql(merged_df, base_file_name)  # Use the base file name as the table name

        download_table_data (base_file_name)
        return send_file(output_file, as_attachment=True)
    else:
        flash('No valid Excel files found!', 'danger')
        return redirect(url_for('index'))

def merge_excel_files(folder_path, folder_name):
    output_file = os.path.join(MERGE_FOLDER, f"{folder_name}.xlsx")  # Output file name matches folder name
    master_df = pd.DataFrame()

    all_files = os.listdir(folder_path)
    xls_files = [file for file in all_files if file.lower().endswith(('.xls', '.xlsx'))]

    if not xls_files:
        return None

    for file in xls_files:
        file_path = os.path.join(folder_path, file)
        try:
            # Read Excel files as HTML tables
            tables = pd.read_html(file_path)
            df = tables[0]

            # Process the data as per your script
            header_row = df.iloc[2]
            data_rows = df.iloc[3:].reset_index(drop=True)
            data_rows.columns = header_row

            new_df = pd.DataFrame()
            new_df['FR CODE'] = data_rows['FR CODE']
            new_df['CONSIGNMENT NUMBER'] = data_rows['CONSIGNMENT NUMBER']
            new_df['MANIFEST NUMBER'] = data_rows['MANIFEST NUMBER']
            new_df['BOOKING DATE'] = data_rows.iloc[:, 3]
            new_df['DESTINATION'] = data_rows.iloc[:, 4]
            new_df['WEIGHT'] = data_rows.iloc[:, 5]
            new_df['CON TYPE'] = data_rows.iloc[:, 6]
            new_df['AMOUNT (Rs.)'] = data_rows.iloc[:, 7]
            new_df['Transhipment'] = data_rows.iloc[:, 8]
            new_df['Service Charge'] = data_rows.iloc[:, 9]
            new_df['Risk Surcharge'] = data_rows.iloc[:, 10]
            new_df['Misc.Charge'] = data_rows.iloc[:, 11]
            new_df['NUMBER OF PIECES'] = data_rows.iloc[:, 12]
            new_df['DESTINATION PINOCDE'] = data_rows.iloc[:, 13]
            new_df['DOX TYPE'] = data_rows.iloc[:, 14]
            new_df['INVOICE NO'] = data_rows.iloc[:, 15]
            new_df['INVOICE DATE'] = data_rows.iloc[:, 16]

            new_df = new_df[new_df['FR CODE'].notna()].drop_duplicates()
            master_df = pd.concat([master_df, new_df], ignore_index=True)
        except Exception as e:
            print(f"Error processing {file}: {e}")
            continue

    if not master_df.empty:
        master_df.to_excel(output_file, index=False)
        return output_file
    return None

def insert_into_postgresql(df, table_name):
    try:
        # Clean column names
        df.columns = df.columns.str.replace(' ', '').str.replace('*', '')

        # Insert data into PostgreSQL
        df.to_sql(table_name, engine, if_exists='append', index=False)
        print(f"✅ Data inserted successfully into the PostgreSQL table '{table_name}'!")
    except Exception as e:
        print(f"Error inserting into PostgreSQL: {e}")

@app.route('/download_table/<table_name>')
def download_table_data(table_name):
    try:
        # Create a direct query to get all data from the table
        query = text(f"""SELECT * ,CASE WHEN "RATE CATEGORIES" IN ('BLUE RATE','GREEN RATE','GOLD RATE','F1 RATE','F2 RATE') THEN 'PERFECT'
				WHEN "DIFF OF GOLD" < 0 THEN 'CREDIT NOTE'
				WHEN "RATE CATEGORIES" = 'ARTO VALUE' THEN 'ARTO VALUE'
				WHEN  "DIFF OF BLUE" IS NOT NULL  OR
			   		  "DIFF OF GREEN"  IS NOT NULL OR 
					  "DIFF OF GOLD" IS NOT NULL OR
					  "DIFF OF F1" IS NOT NULL OR 
					  "DIFF OF F2" IS NOT NULL  THEN 'REVENUE' 
				ELSE NULL  END  AS "PER RATE"			
FROM
(SELECT *  ,
       CASE WHEN "CATEGORIES" = 'P-STANDARD'
	   		THEN (CASE WHEN "DIFF OF BLUE" = 0 AND  
			   				"DIFF OF GREEN" != 0 AND  
							 "DIFF OF GOLD" != 0 AND
							 "DIFF OF F1" != 0  THEN 'BLUE RATE'
		    		   WHEN  "DIFF OF BLUE" != 0 AND  "DIFF OF GREEN" = 0 AND  "DIFF OF GOLD" != 0
	   		              AND "DIFF OF F1" != 0  THEN 'GREEN RATE'
		               WHEN  "DIFF OF BLUE" != 0 AND  "DIFF OF GREEN" != 0 AND  "DIFF OF GOLD" = 0
	   		              AND "DIFF OF F1" != 0  THEN 'GOLD RATE'
		               WHEN  "DIFF OF BLUE" != 0 AND  "DIFF OF GREEN" != 0 AND  "DIFF OF GOLD" != 0
	   		             AND "DIFF OF F1" = 0 THEN 'F1 RATE'
					   ELSE  NULL END )
		    WHEN "CATEGORIES" IN ('D- AIR CARGO','V-PLUS''E-PTP','D- SURFACE')
		    THEN (CASE WHEN  "DIFF OF BLUE" = 0 AND  "DIFF OF GREEN" != 0 AND  "DIFF OF GOLD" != 0
	   		  			  AND "DIFF OF F1" != 0 AND "DIFF OF F2" !=0 THEN 'BLUE RATE'
		    		   WHEN  "DIFF OF BLUE" != 0 AND  "DIFF OF GREEN" = 0 AND  "DIFF OF GOLD" != 0
	   		              AND "DIFF OF F1" != 0 AND "DIFF OF F2" !=0 THEN 'GREEN RATE'
		               WHEN  "DIFF OF BLUE" != 0 AND  "DIFF OF GREEN" != 0 AND  "DIFF OF GOLD" = 0
	   		              AND "DIFF OF F1" != 0 AND "DIFF OF F2" !=0 THEN 'GOLD RATE'
		               WHEN  "DIFF OF BLUE" != 0 AND  "DIFF OF GREEN" != 0 AND  "DIFF OF GOLD" != 0
	   		             AND "DIFF OF F1" = 0 AND "DIFF OF F2" !=0 THEN 'F1 RATE'
			           WHEN  "DIFF OF BLUE" != 0 AND  "DIFF OF GREEN" != 0 AND  "DIFF OF GOLD" != 0
	   		             AND "DIFF OF F1" != 0 AND "DIFF OF F2" = 0 THEN 'F2 RATE' END )
			WHEN "CATEGORIES" = 'ARTO' THEN 'ARTO VALUE'
		               ELSE 'NOT IN CATEGORIE' END AS  "RATE CATEGORIES" 
FROM
	(SELECT DISTINCT "CONSIGNMENTNUMBER","FRCODE","MANIFESTNUMBER","BOOKINGDATE","DESTINATION","WEIGHT","CONTYPE",
			"AMOUNT(Rs.)","Transhipment","ServiceCharge","RiskSurcharge","Misc.Charge","NUMBEROFPIECES",
			"DESTINATIONPINOCDE","DOXTYPE","INVOICENO","INVOICEDATE","CATEGORIES","DESTINATION TYPE",
			"DESTINATIONS","ServiceType","BLUE RATE","BLUE RATE"-"Transhipment" AS "DIFF OF BLUE",
			"GREEN RATE","GREEN RATE" - "Transhipment" AS "DIFF OF GREEN",
			"GOLD RATE", "GOLD RATE" - "Transhipment" AS "DIFF OF GOLD",
			"F1 RATE", "F1 RATE" - "Transhipment" AS "DIFF OF F1",
			"F2 RATE" ,"F2 RATE" - "Transhipment" AS "DIFF OF F2"																																																																													
	FROM
		(SELECT *, CASE WHEN "CATEGORIES" = 'P-STANDARD' 
		    			THEN (CASE WHEN "WEIGHT" < 0.250 THEN "BL-250gms(P)"
        				 			WHEN "WEIGHT" < 0.500 THEN "BL-500gms(P)"
        			     			WHEN CEIL("WEIGHT") > 0.500 THEN 
            			 			"BL-500gms(P)" + ((CEIL(("WEIGHT" - 0.5) / 0.5)) * "BL-Addl.500gms(P)")
        			     			ELSE 10000000000000000  END)
						WHEN "CATEGORIES" = 'D- SURFACE'
						THEN (CASE WHEN "DESTINATIONS"  IN ('Pune City','Mumbai & Maharastra (excluding Vidarbha/Nagpur)',
                 						'Vidarbha/Nagpur, Goa, Gujarat, Madhya Pradesh & Chhattisgarh')
                       				THEN (CASE WHEN "WEIGHT" <= 1 THEN "BL1Kg(DS)"  * "WEIGHT" 
				                  			   WHEN "WEIGHT" <= 2 AND "WEIGHT" > 1 THEN "BL1.01-2Kg(DS)" * "WEIGHT"
				                  		       WHEN "WEIGHT" <= 3 AND "WEIGHT" > 2 THEN "BL2.01-3Kg(DS)" * "WEIGHT"
				                  			   WHEN "WEIGHT" <= 5 THEN "BL<5Kg(DS)" * "WEIGHT"
				                  			   WHEN "WEIGHT" <= 10 THEN "BL<10Kg(DS)" * "WEIGHT"
				                  			   WHEN "WEIGHT" <= 25 THEN "BL<25Kg(DS)" * "WEIGHT"
				                               WHEN "WEIGHT" <= 50 THEN "BL<50Kg(DS)" * "WEIGHT"
				                  	           WHEN "WEIGHT" > 50 THEN  "BL<50Kg(DS)" * "WEIGHT" END ) 
			            	       WHEN "DESTINATIONS"  IN ('Metros (Kolkata, Delhi, Bangalore, Hyderabad, & Chennai) & Cochin',
				                        'East, South & North Zone (except North East, Port Blair, Jammu & Kashmir, Himachal Pradesh & Leh)',
                                        'Jammu & Kashmir, Himachal Pradesh & Gauhati','North East (including Tripura)')
				                  THEN (CASE WHEN "WEIGHT" <= 5 THEN  "BL<5Kg(DS)"  * 5
							                 WHEN "WEIGHT" <=10 THEN  "BL<10Kg(DS)"  * "WEIGHT"
							                 WHEN "WEIGHT" <=25 THEN "BL<25Kg(DS)" * "WEIGHT"
							                 WHEN "WEIGHT" <=50 THEN "BL<50Kg(DS)" * "WEIGHT"
							                 WHEN  "WEIGHT" > 50 THEN "BL>50Kg(DS)"  * "WEIGHT"  END)
				                             ELSE  10000000000 END)
			                      WHEN "CATEGORIES" = 'D- AIR CARGO'
			                      THEN (CASE WHEN "WEIGHT" <= 5 THEN "BL<5Kg(DA)" * "WEIGHT"
				                             WHEN "WEIGHT" <= 10 THEN "BL<10Kg(DA)" * "WEIGHT"
				                             WHEN "WEIGHT" <= 25 THEN "BL<25Kg(DA)" * "WEIGHT"
				                             WHEN "WEIGHT" <= 50 THEN "BL<50Kg(DA)" * "WEIGHT"
				                             WHEN "WEIGHT" > 50 THEN  "BL<50Kg(DA)" * "WEIGHT" 
				                             ELSE  10000000000 END)
			                      ELSE NULL END  AS "BLUE RATE", ------------------------------------ "BLUE RATE"
		                         CASE WHEN "CATEGORIES" = 'P-STANDARD' 
		   						  THEN (CASE  WHEN "WEIGHT" < 0.250 THEN "GRN-250gms(P)"
          				 					  WHEN "WEIGHT" < 0.500 THEN "GRN-500gms(P)"
          				  	                  WHEN CEIL("WEIGHT") > 0.500 THEN 
            				                  "GRN-500gms(P)" + ((CEIL(("WEIGHT" - 0.5) / 0.5)) * "GRN-Addl.500gms(P)")
          				                      ELSE 10000000000000000 END)
								  WHEN "CATEGORIES" = 'D- SURFACE'
								  THEN (CASE WHEN "DESTINATIONS"  IN ('Pune City','Mumbai & Maharastra (excluding Vidarbha/Nagpur)',
					  					      'Vidarbha/Nagpur, Goa, Gujarat, Madhya Pradesh & Chhattisgarh')
            					  		     THEN (CASE WHEN "WEIGHT" <= 1 THEN "GRN0-1Kg(DS)" * "WEIGHT"
				 					                   WHEN "WEIGHT" <= 2 AND "WEIGHT" > 1 THEN "GRN1.01-2Kg(DS)" * "WEIGHT"
				 					  				   WHEN "WEIGHT" <= 3 AND "WEIGHT" > 2 THEN "GRN2.01-3Kg(DS)" * "WEIGHT"
									  				   WHEN "WEIGHT" <= 5 THEN "GRN<5Kg(DS)" * "WEIGHT"
				 					                   WHEN "WEIGHT" <= 10 THEN "GRN<10Kg(DS)" * "WEIGHT"
				                                       WHEN "WEIGHT" <= 25 THEN "GRN<25Kg(DS)" * "WEIGHT"
				                                       WHEN "WEIGHT" <= 50 THEN "GRN<50Kg(DS)" * "WEIGHT"
				                                       WHEN "WEIGHT" > 50 THEN  "GRN<50Kg(DS)" * "WEIGHT" END) 
			                     WHEN "DESTINATIONS"  IN ('Metros (Kolkata, Delhi, Bangalore, Hyderabad, & Chennai) & Cochin',
					                                      'East, South & North Zone (except North East, Port Blair, Jammu & Kashmir, Himachal Pradesh & Leh)',
                                                          'Jammu & Kashmir, Himachal Pradesh & Gauhati','North East (including Tripura)')
				                 THEN  (CASE WHEN "WEIGHT" <= 5 THEN  "GRN<5Kg(DS)" * 5
							                WHEN "WEIGHT" <=10 THEN  "GRN<10Kg(DS)" * "WEIGHT"
							                WHEN "WEIGHT" <=25 THEN "GRN<25Kg(DS)" * "WEIGHT"
							                WHEN "WEIGHT" <=50 THEN "GRN<50Kg(DS)" * "WEIGHT"
							                WHEN  "WEIGHT" > 50 THEN "GRN>50Kg(DS)" * "WEIGHT" END)
				                 ELSE  10000000000 END) 
			                     WHEN "CATEGORIES" = 'D- AIR CARGO'
			                     THEN (CASE  WHEN "WEIGHT" <= 10 THEN "GRN<10Kg(DA)" * "WEIGHT"
					                         WHEN "WEIGHT" <= 25 THEN "GRN<25Kg(DA)" * "WEIGHT"
					                         WHEN "WEIGHT" <= 50 THEN "GRN<50Kg(DA)" * "WEIGHT"
					                         WHEN "WEIGHT" > 50 THEN  "GRN<50Kg(DA)" * "WEIGHT"
					                   ELSE  10000000000 END)
			                     ELSE NULL END  AS "GREEN RATE", ------------------------------------ "GREEN RATE"	
		                         CASE WHEN "CATEGORIES" = 'P-STANDARD' 
		                         THEN (CASE WHEN "WEIGHT" < 0.250 THEN "GLD-250gms(P)"
                                            WHEN "WEIGHT" < 0.500 THEN "GLD-500gms(P)"
          				                    WHEN CEIL("WEIGHT") > 0.500 THEN 
                                            "GLD-500gms(P)" + ((CEIL(("WEIGHT" - 0.5) / 0.5)) * "GLD-Addl.500gms(P)")
          				                    ELSE 10000000000000000 END)
			                   WHEN "CATEGORIES" = 'D- SURFACE'
			THEN (CASE WHEN "DESTINATIONS"  IN ('Pune City','Mumbai & Maharastra (excluding Vidarbha/Nagpur)',
'Vidarbha/Nagpur, Goa, Gujarat, Madhya Pradesh & Chhattisgarh')
            THEN CASE WHEN "WEIGHT" <= 1 THEN "GLD0-1Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 2 AND "WEIGHT" > 1 THEN "GLD1.01-2Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 3 AND "WEIGHT" > 2 THEN "GLD2.01-3Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 5 THEN "GLD<5Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 10 THEN "GLD<10Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 25 THEN "GLD<25Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 50 THEN "GLD<50Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" > 50 THEN  "GLD<50Kg(DS)" * "WEIGHT" END  
			 WHEN "DESTINATIONS"  IN ('Metros (Kolkata, Delhi, Bangalore, Hyderabad, & Chennai) & Cochin',
					'East, South & North Zone (except North East, Port Blair, Jammu & Kashmir, Himachal Pradesh & Leh)',
                    'Jammu & Kashmir, Himachal Pradesh & Gauhati','North East (including Tripura)')
				THEN  CASE WHEN "WEIGHT" <= 5 THEN  "GLD<5Kg(DS)" * 5 
							WHEN "WEIGHT" <=10 THEN  "GLD<10Kg(DS)"  * "WEIGHT"
							WHEN "WEIGHT" <=25 THEN "GLD<25Kg(DS)" * "WEIGHT"
							WHEN "WEIGHT" <=50 THEN "GLD<50Kg(DS)" * "WEIGHT"
							WHEN  "WEIGHT" > 50 THEN "GLD>50Kg(DS)" * "WEIGHT" END 
				ELSE  10000000000 END)
			WHEN "CATEGORIES" = 'D- AIR CARGO'
				THEN( CASE WHEN "WEIGHT" <= 5 THEN  "GLD<5Kg(DA)" * 5 
					WHEN "WEIGHT" <=10 THEN  "GLD<10Kg(DA)"  * "WEIGHT"
					WHEN "WEIGHT" <=25 THEN "GLD<25Kg(DA)" * "WEIGHT"
					WHEN "WEIGHT" <=50 THEN "GLD<50Kg(DA)" * "WEIGHT"
					WHEN  "WEIGHT" > 50 THEN "GLD>50Kg(DA)" * "WEIGHT" 
					ELSE  10000000000 END)
			ELSE NULL END  AS "GOLD RATE", ------------------------------------------"GOLD RATE"
        CASE WHEN "CATEGORIES" = 'P-STANDARD' 
		    THEN (CASE WHEN "WEIGHT" < 0.250 THEN "F1-250gms(P)"
          			     WHEN "WEIGHT" < 0.500 THEN "F1-500gms(P)"
                         WHEN CEIL("WEIGHT") > 0.500 THEN 
                         "F1-500gms(P)" + ((CEIL(("WEIGHT" - 0.5) / 0.5)) * "F1-Addl.500gms(P)")
          				  ELSE 10000000000000000  END)
			WHEN "CATEGORIES" = 'D- SURFACE'
			THEN (CASE WHEN "DESTINATIONS"  IN ('Pune City','Mumbai & Maharastra (excluding Vidarbha/Nagpur)',
'Vidarbha/Nagpur, Goa, Gujarat, Madhya Pradesh & Chhattisgarh')
            THEN CASE WHEN "WEIGHT" <= 1 THEN "F10-1Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 2 AND "WEIGHT" > 1 THEN "F11.01-2Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 3 AND "WEIGHT" > 2 THEN "F12.01-3Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 5 THEN "F1<5Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 10 THEN "F1<10Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 25 THEN "F1<25Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 50 THEN "F1<50Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" > 50 THEN  "F1<50Kg(DS)" * "WEIGHT" END  
			 WHEN "DESTINATIONS"  IN ('Metros (Kolkata, Delhi, Bangalore, Hyderabad, & Chennai) & Cochin',
					'East, South & North Zone (except North East, Port Blair, Jammu & Kashmir, Himachal Pradesh & Leh)',
                    'Jammu & Kashmir, Himachal Pradesh & Gauhati','North East (including Tripura)')
				THEN  CASE WHEN "WEIGHT" <= 5 THEN  "F1<5Kg(DS)"  * 5
							WHEN "WEIGHT" <=10 THEN  "F1<10Kg(DS)"  * "WEIGHT"
							WHEN "WEIGHT" <=25 THEN "F1<25Kg(DS)" * "WEIGHT"
							WHEN "WEIGHT" <=50 THEN "F1<50Kg(DS)" * "WEIGHT"
							WHEN  "WEIGHT" > 50 THEN "F1>50Kg(DS)" * "WEIGHT" END 
				ELSE  10000000000 END) 
			WHEN "CATEGORIES" = 'D- AIR CARGO'
			THEN (CASE WHEN "WEIGHT" <= 5 THEN  "F1<5Kg(DA)"  * 5
								WHEN "WEIGHT" <=10 THEN  "F1<10Kg(DA)"  * "WEIGHT"
								WHEN "WEIGHT" <=25 THEN "F1<25Kg(DA)" * "WEIGHT"
								WHEN "WEIGHT" <=50 THEN "F1<50Kg(DA)" * "WEIGHT"
								WHEN  "WEIGHT" > 50 THEN "F1>50Kg(DA)" * "WEIGHT"  
					ELSE  10000000000 END)
			ELSE NULL END  AS "F1 RATE",-----------------------------------------------"F1 RATE"
			
		 CASE  --WHEN "CATEGORIES" = 'P-STANDARD' 
		--     THEN 10000000000
			WHEN "CATEGORIES" = 'D- SURFACE'
			THEN (CASE WHEN "DESTINATIONS"  IN ('Pune City','Mumbai & Maharastra (excluding Vidarbha/Nagpur)',
'Vidarbha/Nagpur, Goa, Gujarat, Madhya Pradesh & Chhattisgarh')
            THEN CASE WHEN "WEIGHT" <= 1 THEN "F20-1Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 2 AND "WEIGHT" > 1 THEN "F21.01-2Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 3 AND "WEIGHT" > 2 THEN "F22.01-3Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 5 THEN "F2<5Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 10 THEN "F2<10Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 25 THEN "F2<25Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" <= 50 THEN "F2<50Kg(DS)" * "WEIGHT"
				 WHEN "WEIGHT" > 50 THEN  "F2<50Kg(DS)"  * "WEIGHT" END  
			 WHEN "DESTINATIONS"  IN ('Metros (Kolkata, Delhi, Bangalore, Hyderabad, & Chennai) & Cochin',
					'East, South & North Zone (except North East, Port Blair, Jammu & Kashmir, Himachal Pradesh & Leh)',
                    'Jammu & Kashmir, Himachal Pradesh & Gauhati','North East (including Tripura)')
				THEN  CASE WHEN "WEIGHT" <= 5 THEN  "F2<5Kg(DS)"  * 5
							WHEN "WEIGHT" <=10 THEN  "F2<10Kg(DS)"  * "WEIGHT"
							WHEN "WEIGHT" <=25 THEN "F2<25Kg(DS)" * "WEIGHT"
							WHEN "WEIGHT" <=50 THEN "F2<50Kg(DS)" * "WEIGHT"
							WHEN  "WEIGHT" > 50 THEN "F2>50Kg(DS)"  * "WEIGHT" END 
				ELSE  10000000000 END) 
              WHEN "CATEGORIES" = 'D- AIR CARGO'
			  THEN (CASE WHEN "WEIGHT" <= 5 THEN  "F2<5Kg(DA)"  * 5
								WHEN "WEIGHT" <=10 THEN  "F2<10Kg(DA)"  * "WEIGHT"
								WHEN "WEIGHT" <=25 THEN "F2<25Kg(DA)" * "WEIGHT"
								WHEN "WEIGHT" <=50 THEN "F2<50Kg(DA)" * "WEIGHT"
								WHEN  "WEIGHT" > 50 THEN "F2>50Kg(DA)"  * "WEIGHT" 
					ELSE  10000000000 END)
			ELSE NULL END  AS "F2 RATE" ------------------------------------------"F1 RATE"		
FROM
(SELECT * FROM 
					(SELECT *,CASE WHEN A."CATEGORIES" = 'D- AIR CARGO' THEN "DAIRCARGO"  
               					   WHEN A."CATEGORIES" = 'P-STANDARD' THEN "PSIRES"
                				   WHEN A."CATEGORIES" = 'V-PLUS' THEN "VPLUS"
                				   WHEN A."CATEGORIES" = 'D- SURFACE' THEN "DSURFACECARGO"
								   WHEN A."CATEGORIES" = 'ARTO' THEN 'ARTO VALUE'
								   ELSE NULL END AS "DESTINATION TYPE"
					  FROM 
                           (SELECT *, CASE When "CONTYPE" in ('AC1','AC') Then 'D- AIR CARGO'
				 							When SUBSTRING("CONSIGNMENTNUMBER" FROM 1 FOR 1)= '0' Then 'ARTO'
	       		 							When SUBSTRING("CONSIGNMENTNUMBER" FROM 1 FOR 1)= 'P' Then 'P-STANDARD'
		   		 							WHEN SUBSTRING("CONSIGNMENTNUMBER" FROM 1 FOR 1)= 'V' Then 'V-PLUS'
		   		 							When  SUBSTRING("CONSIGNMENTNUMBER" FROM 1 FOR 1)= 'E' Then 'E-PTP'
		   		 							When "CONTYPE" IN ('SF1','SF')  Then 'D- SURFACE'
		  		 							ELSE 'NULL' END AS "CATEGORIES" 
    						FROM {table_name})A 
					  LEFT JOIN PINCODE B 
					  ON A."DESTINATIONPINOCDE" = B."PINCODE") C
LEFT JOIN "All Service" D
ON  C."DESTINATION TYPE" = D."DESTINATIONS" AND  C."CATEGORIES" = D."ServiceType"))))
 """)

        # Read the query results directly into a DataFrame
        df = pd.read_sql(query, engine)

        # Create a file path for the Excel export
        DB_FILE = 'DataBase'
        output_file = os.path.join(DB_FILE, f"{table_name}_database.xlsx")

        # Save DataFrame to Excel
        df.to_excel(output_file, index=False)

        # Send the file as an attachment
        return send_file(output_file, as_attachment=True, download_name=f"{table_name}_database.xlsx")

    except Exception as e:
        print(f"Error executing SELECT query: {e}")
        flash(f'Error retrieving data from table {table_name}', 'danger')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)