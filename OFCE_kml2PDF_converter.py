import os
import sys
import webbrowser
import xml.etree.ElementTree as ET
from reportlab.pdfgen import canvas
from datetime import datetime
from reportlab.lib.pagesizes import letter
from datetime import timedelta


def parse_kml(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    placemarks = root.findall(".//{http://www.opengis.net/kml/2.2}Placemark")

    data = []
    for placemark in placemarks:
        name = placemark.find("{http://www.opengis.net/kml/2.2}name").text
        coordinates = placemark.find(".//{http://www.opengis.net/kml/2.2}coordinates").text
        timestamp_start = placemark.find(".//{http://www.opengis.net/kml/2.2}TimeSpan/{http://www.opengis.net/kml/2.2}begin").text
        timestamp_end = placemark.find(".//{http://www.opengis.net/kml/2.2}TimeSpan/{http://www.opengis.net/kml/2.2}end").text
        description = placemark.find("{http://www.opengis.net/kml/2.2}description").text if placemark.find(
            "{http://www.opengis.net/kml/2.2}description") is not None else ""
        address = placemark.find("{http://www.opengis.net/kml/2.2}address").text if placemark.find(
            "{http://www.opengis.net/kml/2.2}address") is not None else ""

        data.append({
            'name': name,
            'coordinates': coordinates,
            'timestamp_start': timestamp_start,
            'timestamp_end': timestamp_end,
            'description': description,
            'address': address
        })

    return data

def create_pdf(data, input_filename, start_date=None, end_date=None):
    base_filename = os.path.splitext(os.path.basename(input_filename))[0]

    # Replace invalid characters with underscores
    sanitized_email = base_filename.replace('@', '_').replace('.', '_')


    if start_date and end_date and start_date == end_date:
        pdf_filename = f"{sanitized_email}_GPS_{start_date.replace('/', '_')}.pdf"
    elif start_date and end_date:
        pdf_filename = f"{sanitized_email}_GPS_{start_date.replace('/', '_')}-{end_date.replace('/', '_')}.pdf"
    else:
        pdf_filename = f"{sanitized_email}_GPS_complete.pdf"

    c = canvas.Canvas(pdf_filename, pagesize=letter)

    c.setFont("Helvetica-Bold", 20)
    title1 = "Location History Report"
    title2 = f"File: {base_filename}.kml"
    title3 = f"Start Date: {start_date}" if start_date else ""
    title4 = f"End Date: {end_date}" if start_date else ""
    title5 = f"Date: {start_date}"
    title1_width = c.stringWidth(title1, "Helvetica", 20)
    title2_width = c.stringWidth(title2, "Helvetica", 20)
    title3_width = c.stringWidth(title3, "Helvetica", 12)
    title_total_width = max(title1_width, title2_width, title3_width)

    c.drawCentredString(letter[0] / 2, letter[1] / 2 + 40, title1)

    # Add an empty line
    c.drawCentredString(letter[0] / 2, letter[1] / 2 - 20, "")

    c.drawCentredString(letter[0] / 2, letter[1] / 2, title2)

    if (start_date and end_date and start_date == end_date):
        c.drawCentredString(letter[0] / 2, letter[1] / 2 - 40, title5)
    if not (start_date and end_date and start_date == end_date):
        c.drawCentredString(letter[0] / 2, letter[1] / 2 - 40, title3)
        c.drawCentredString(letter[0] / 2, letter[1] / 2 - 60, title4)

    # Add "Complete report" if start_date and end_date are not provided
    if not start_date and not end_date and data:
        start_date = datetime.strptime(data[0]['timestamp_start'], "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%d/%m/%Y")
        end_date = datetime.strptime(data[-1]['timestamp_end'], "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%d/%m/%Y")
        location_date = f"Locations from {start_date} to {end_date}"
        c.setFont("Helvetica-Bold", 18)
        c.drawCentredString(letter[0] / 2, letter[1] / 2 - 80, "Complete report")
        c.setFont("Helvetica", 16)
        c.drawCentredString(letter[0] / 2, letter[1] / 2 - 110, location_date)

    c.showPage()
    c.setFont("Helvetica", 12)

    top_margin = 10
    bottom_margin = 10
    current_day = None
    vertical_position = 750

    for entry in data:
        name = entry['name']
        description = entry['description']
        coordinates = entry['coordinates']
        timestamp_start = entry['timestamp_start']
        timestamp_end = entry['timestamp_end']
        address = entry['address']

        day_start = datetime.strptime(timestamp_start, "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%d/%m/%Y")
        formatted_timestamp_start = datetime.strptime(timestamp_start,
                                                       "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%Y-%m-%d %H:%M:%S")

        day_end = datetime.strptime(timestamp_end, "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%Y-%m-%d")
        formatted_timestamp_end = datetime.strptime(timestamp_end,
                                                     "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%Y-%m-%d %H:%M:%S")

        # Check if a new day is encountered
        if day_start != current_day:
            if current_day is not None:
                vertical_position = 750 - top_margin  # Reset vertical position for the new page (with top margin)
            current_day = day_start

            # Add a page with the current day in large font
            c.setFont("Helvetica-Bold", 20)
            formatted_day_start = datetime.strptime(day_start, "%d/%m/%Y").strftime("%A %d/%m/%Y")
            c.drawCentredString(letter[0] / 2, letter[1] / 2, formatted_day_start)
            c.showPage()  # Move to the next page for the regular content
            c.setFont("Helvetica", 12)

            vertical_position = 750 - top_margin  # Reset the vertical position for the new page (with top margin)

        c.setFont("Helvetica-Bold", 12)
        c.drawString(100, vertical_position, f"Name: {name}")
        vertical_position -= 10  # Adjust vertical position for the next line

        if address:
            c.drawString(100, vertical_position - 20, f"Address: {address}")
            vertical_position -= 20  # Adjust vertical position for the next line

        c.setFont("Helvetica", 12)  # Reset font to regular

        if "Distance" in description:
            distance_index = description.find("Distance")
            first_part = description[:distance_index].strip()
            second_part = description[distance_index + len("Distance"):].strip()

            c.setFont("Helvetica-Bold", 12)
            c.drawString(100, vertical_position - 40, "Description:")
            c.setFont("Helvetica", 12)
            description_x = 100 + c.stringWidth("Description:", "Helvetica-Bold", 12) + 10
            c.drawString(description_x, vertical_position - 40, f"{first_part}")

            c.setFont("Helvetica-Bold", 12)
            c.drawString(100, vertical_position - 60, "Distance:")
            c.setFont("Helvetica", 12)
            distance_x = 100 + c.stringWidth("Distance:", "Helvetica-Bold", 12) + 10
            c.drawString(distance_x, vertical_position - 60, f"{second_part}")

        else:
            c.setFont("Helvetica-Bold", 12)
            c.drawString(100, vertical_position - 40, "Description:")
            c.setFont("Helvetica", 12)
            description_x = 100 + c.stringWidth("Description:", "Helvetica-Bold", 12) + 10
            c.drawString(description_x, vertical_position - 40, f"{description}")

        c.setFont("Helvetica-Bold", 12)
        c.drawString(100, vertical_position - 80, "Timestamp start:")
        c.setFont("Helvetica", 12)
        c.drawString(100 + c.stringWidth("Timestamp start:", "Helvetica-Bold", 12) + 10, vertical_position - 80,
                     formatted_timestamp_start)

        c.setFont("Helvetica-Bold", 12)
        c.drawString(100, vertical_position - 100, "Timestamp end:")
        c.setFont("Helvetica", 12)
        c.drawString(100 + c.stringWidth("Timestamp end:", "Helvetica-Bold", 12) + 10, vertical_position - 100,
                     formatted_timestamp_end)

        c.setFont("Helvetica-Bold", 12)
        c.drawString(100, vertical_position - 120, "Coordinates:")
        c.setFont("Helvetica", 12)

        coordinates_label_width = c.stringWidth("Coordinates:", "Helvetica-Bold", 12)
        coordinates_lines = [f"{pair.strip()}" for pair in coordinates.split(' ') if pair.strip()]

        for line in coordinates_lines:
            coordinates_x = 100 + coordinates_label_width + 10

            if vertical_position - 20 < 120:
                c.showPage()
                vertical_position = 800
                c.setFont("Helvetica", 12)

            c.drawString(coordinates_x, vertical_position - 120, line)
            vertical_position -= 20

        vertical_position = 750 - top_margin
        c.showPage()

    c.save()
    print(f"PDF file '{pdf_filename}' created successfully.")
    return pdf_filename

def filter_entries_by_date(entries, start_date, end_date):
    filtered_entries = []
    for entry in entries:
        timestamp_start = datetime.strptime(entry['timestamp_start'], "%Y-%m-%dT%H:%M:%S.%fZ")
        timestamp_end = datetime.strptime(entry['timestamp_end'], "%Y-%m-%dT%H:%M:%S.%fZ")

# Consider entries with a start or end timestamp within the range of start_date and end_date
# Add 1 day to end_date to include the entire day

        if start_date <= timestamp_start <= end_date or start_date <= timestamp_end <= (end_date + timedelta(days=1)):
            filtered_entries.append(entry)
    return filtered_entries


def print_manual_dates(data, input_filename):
    created_files = []  # Lista per tenere traccia dei file PDF creati

    while True:
        print("\nEnter the date to print (dd/mm/yyyy) or 'q' to quit:")
        date_input = input()

        if date_input.lower() == 'q':
            break

        try:
            selected_date = datetime.strptime(date_input, "%d/%m/%Y")
            selected_entries = filter_entries_by_date(data, selected_date, selected_date)

            if selected_entries:
                pdf_filename = create_pdf(selected_entries, input_filename, date_input, date_input)
                created_files.append(pdf_filename) 
                print(f"PDF file for {date_input} created successfully.")
            else:
                print(f"No entries found for {date_input}.")
        except ValueError:
            print("Invalid date format. Please use the format dd/mm/yyyy.")

    
    for pdf_file in created_files:
        webbrowser.open(pdf_file)



def main():


    print("""
    Welcome to the OFCE KML to PDF converter!

    This script allows you to create PDF reports from KML files produced by Oxygen Forensics Cloud Extractor containing location history data. Follow the steps below to get started:

    1. Place your KML files in the same directory as this script.
    2. Run the script and select a KML file from the list.
    3. Choose an option to generate PDF reports:
        - Option 1: Print all entries.
        - Option 2: Print a selection based on dates.
        - Option 3: Print manually selected dates.

    4. Follow the prompts to customize your report based on your chosen option.
    5. After each operation, decide whether to perform another action.

    Note: To exit the script, type 'n' when asked if you want to perform another operation.

    Let's get started!
    """)

    created_files = []  # List to keep track of created PDF files

    while True:
        kml_files = [f for f in os.listdir() if f.endswith(".kml")]

        if not kml_files:
            print("No KML files found in the current directory.")
            return

        print("Select a KML file:")
        for i, file in enumerate(kml_files, 1):
            print(f"{i}. {file}")

        selected_index = int(input("Enter the number of the KML file: ")) - 1

        try:
            selected_file = kml_files[selected_index]
            kml_data = parse_kml(selected_file)

            while True:
                print("Select an option:")
                print("1. Print all entries")
                print("2. Print a selection based on dates")
                print("3. Print manually selected dates")

                choice = int(input("Enter your choice (1, 2, or 3): "))

                if choice == 1:
                    pdf_filename = create_pdf(kml_data, selected_file)
                    created_files.append(pdf_filename)
                elif choice == 2:
                    start_date_str = input("Enter the start date (dd/mm/yyyy): ")
                    end_date_str = input("Enter the end date (dd/mm/yyyy): ")

                    start_date = datetime.strptime(start_date_str, "%d/%m/%Y")
                    end_date = datetime.strptime(end_date_str, "%d/%m/%Y")

                    selected_entries = filter_entries_by_date(kml_data, start_date, end_date)

                    if selected_entries:
                        pdf_filename = create_pdf(selected_entries, selected_file, start_date_str, end_date_str)
                        created_files.append(pdf_filename)
                        print("PDF file created successfully.")
                    else:
                        print("No entries found within the specified date range.")
                elif choice == 3:
                    print_manual_dates(kml_data, selected_file)
                else:
                    print("Invalid choice. Please enter 1, 2, or 3.")
                    continue  # Continue to the next iteration of the inner loop

                # Ask the user if they want to perform another operation
                another_request = input("Do you want to perform another operation? (y/n): ")
                if another_request.lower() != 'y':
                    print("Exiting the program. Goodbye!")
                    for pdf_file in created_files:
                        webbrowser.open(pdf_file)
                    sys.exit()  # Termina lo script

        except IndexError:
            print("Invalid selection. Please enter a valid number.")
        except ValueError as e:
            print(f"Error: {e}")

            print("Exiting the program. Goodbye!")
            for pdf_file in created_files:
                webbrowser.open(pdf_file)
            sys.exit()  # Termina lo script

if __name__ == "__main__":
    main()
