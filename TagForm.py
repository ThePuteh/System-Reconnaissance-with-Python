import os
import platform
import csv
import subprocess
import socket
import time

def display_welcome_message():
    print("JPN Reconnaisance [Version 1.00]")
    print("(c) Puteh the Analyst. All rights reserved.\nFollow me on Github at https://github.com/ThePuteh")
    print()

def clear_screen():
    os.system('cls')

def get_equipment_type():
    while True:
        try:
            print("Choose Equipment Type:\n")
            print("1) Desktop")
            print("2) Laptop\n")
            choice = input("Enter your choice (1 or 2): ").upper()
            clear_screen()

            if choice == '1':
                return 'Desktop'
            elif choice == '2':
                return 'Laptop'
            else:
                print("Invalid choice. Please enter 1 or 2.")

        except Exception as e:
            print(f"Error: {e}")

def get_model():
    try:
        wmic_result = subprocess.run(['wmic', 'csproduct', 'get', 'name'], capture_output=True, text=True, check=True)
        model = wmic_result.stdout.strip().split('\n')[-1].strip().upper()
        return model
    except Exception as e:
        print(f"Error: {e}")
        return None

def get_os_info():
    try:
        os_name = platform.system()
        os_version = platform.version()
        os_info = f"{os_name} {os_version}".upper()
        return os_info
    except Exception as e:
        print(f"Error: {e}")
        return None

def get_serial_number():
    try:
        result = subprocess.run(['wmic', 'bios', 'get', 'serialnumber'], capture_output=True, text=True, check=True)
        serial_number = result.stdout.strip().split('\n')[-1].strip().upper()
        return serial_number
    except subprocess.CalledProcessError as e:
        print(f"Error while getting Serial Number: {e}")
        return None

def get_kewpa_number():
    while True:
        try:
            has_kewpa = input("Do you have a KEWPA No.? (y/n): ").upper()
            if has_kewpa == 'Y':
                first_data = input("Enter the SECOND LAST number of KEWPA No.: ")
                last_data = input("Enter the LAST number of KEWPA No.: ")
                if first_data.isdigit() and last_data.isdigit():
                    return f"{first_data.strip()} / {last_data.strip()}"
                else:
                    print("Invalid input. Please enter numeric values for KEWPA No.")
            elif has_kewpa == 'N':
                return "Not Available"
            else:
                print("Invalid choice. Please enter 'Y' or 'N'.")
        except Exception as e:
            print(f"Error: {e}")

def get_pc_name():
    try:
        pc_name = platform.node()
        return pc_name
    except Exception as e:
        print(f"Error: {e}")
        return None

def get_ip_address():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip_address = s.getsockname()[0].upper()
        s.close()
        return ip_address
    except Exception as e:
        print(f"Error while getting IP address: {e}")
        return None

def remove_duplicates_from_csv(filename):
    try:
        with open(filename, 'r') as csvfile:
            reader = csv.reader(csvfile)
            rows = list(reader)

        header = rows[0]
        unique_rows = [header]

        seen_serial_numbers = set()
        for row in rows[1:]:
            serial_number = row[header.index('Serial Number')]
            if serial_number not in seen_serial_numbers:
                unique_rows.append(row)
                seen_serial_numbers.add(serial_number)

        with open(filename, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC)
            writer.writerows(unique_rows)
        print("Duplicates removed from CSV.")
    except Exception as e:
        print(f"Error while removing duplicates: {e}")

def save_to_csv(equipment_type, model, os_info, serial_number, kewpa_number, pc_name, ip_address, filename):
    try:
        with open(filename, 'a', newline='') as csvfile:
            csvwriter = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC)
            if csvfile.tell() == 0:
                csvwriter.writerow(['Equipment Type', 'Model', 'Operating System', 'Serial Number', 'Tag No.', 'KEWPA No.', 'PC Name', 'IP Address', "Inteksoft's Remarks", 'Owner'])
            owner_name = input("Enter the Owner's Name: ").strip().upper()

            # Prompt the user to enter the owner's name again if left blank
            while not owner_name:
                print("Owner's Name cannot be left blank.")
                owner_name = input("Enter the Owner's Name: ").strip().upper()

            confirm_save = input(f"Do you want to save the data for {owner_name}? (y/n): ").upper()
            if confirm_save == 'Y':
                csvwriter.writerow([equipment_type, model, os_info, serial_number, '', f"'{kewpa_number}", pc_name, ip_address, '', owner_name])
                print(f"Data for {owner_name} saved to {filename} excel !")
                print("\nRestarting in 3 Second")
                time.sleep(3)
                clear_screen()
            else:
                print("Data not saved. Exiting.")

    except Exception as e:
        print(f"Error while saving to CSV: {e}")

def delete_row_by_serial_number(serial_number, filename):
    try:
        with open(filename, 'r') as csvfile:
            reader = csv.reader(csvfile)
            rows = list(reader)

        header = rows[0]
        found = False
        for row in rows[1:]:
            if row[header.index('Serial Number')] == serial_number:
                found = True
                print("\n===========================================================================")
                print("Details to be deleted:\n")
                print(f"Equipment Type: {row[header.index('Equipment Type')]}")
                print(f"Model: {row[header.index('Model')]}")
                print(f"Operating System: {row[header.index('Operating System')]}")
                print(f"Serial Number: {row[header.index('Serial Number')]}")
                print(f"Tag No.: {row[header.index('Tag No.')]}")
                print(f"KEWPA No.: {row[header.index('KEWPA No.')]}")
                print(f"PC Name: {row[header.index('PC Name')]}")
                print(f"IP Address: {row[header.index('IP Address')]}")
                print(f"Inteksoft's Remarks: {row[header.index('Inteksoft\'s Remarks')]}")
                print(f"Owner: {row[header.index('Owner')]}")
                print("===========================================================================\n")

                confirm_delete = input("Do you want to delete this row? (y/n): ").upper()
                if confirm_delete == 'Y':
                    rows.remove(row)
                    print("Row deleted.")
                else:
                    print("Deletion cancelled.")

        if not found:
            print(f"No row found with Serial Number: {serial_number}")

        with open(filename, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC)
            writer.writerows(rows)
    except Exception as e:
        print(f"Error while deleting row: {e}")

def kbhit():
    try:
        import msvcrt
        return msvcrt.kbhit()
    except ImportError:
        return False

if __name__ == "__main__":
    display_welcome_message()

    while True:
        operation_choice = input("Enter any key :\n(a) for Add Data, \n(d) for Delete Data, \n(r) for Remove Duplicate Serialnumber, \n(x) for Exit?\n\nYou select : \t ").upper()

        if operation_choice == 'A':
            equipment_type = get_equipment_type()
            model = get_model()
            os_info = get_os_info()
            serial_number = get_serial_number()
            kewpa_number = get_kewpa_number()
            pc_name = get_pc_name()
            ip_address = get_ip_address()

            if equipment_type and model and os_info and serial_number and kewpa_number is not None and pc_name and ip_address:
                print("\n===========================================================================\n")
                print("PLEASE CHECK YOUR ENTERED INFO BELOW BEFORE CONFIRMATION")
                print("THE DATA ENTERED GOING TO BE ADD NEXT AS BELOW :\n")
                print(f"Equipment Type: {equipment_type}")
                print(f"Model: {model}")
                print(f"Operating System: {os_info}")
                print(f"Serial Number: {serial_number}")
                print(f"Tag No.: ")
                print(f"KEWPA No.: {kewpa_number}")
                print(f"PC Name: {pc_name}")
                print(f"IP Address: {ip_address}\n")
                print("===========================================================================\n")
                save_to_csv(equipment_type, model, os_info, serial_number, kewpa_number, pc_name, ip_address, 'pc_info.csv')
            else:
                print("Unable to retrieve PC information.")
        elif operation_choice == 'D':
            serial_number_to_delete = input("Enter the Serial Number of the row you want to delete: ")
            delete_row_by_serial_number(serial_number_to_delete, 'pc_info.csv')
        elif operation_choice == 'R':
            remove_duplicates_from_csv('pc_info.csv')
        elif operation_choice == 'X':
            print("Exiting the program in 5 seconds...")
            for i in range(5, 0, -1):
                if kbhit():
                    break
                print(f"{i}", end='\r')
                time.sleep(1)
            break
        else:
            clear_screen()
            print("Invalid choice. Please enter 'A' for add, 'D' for delete, 'R' for removing duplicates, or 'X' to exit.")
    clear_screen()
    print("\nGoodbye!")
