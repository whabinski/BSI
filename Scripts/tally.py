import os

# Function to read the current tally from the file
def read_tally(tally_file='../tally.txt'):
    if os.path.exists(tally_file):
        with open(tally_file, 'r') as file:
            try:
                tally = file.read().strip()  # Read as string
                if tally.isdigit():  # Check if the tally is a valid number
                    tally = tally.zfill(2)  # Ensure two digits
                else:
                    tally = '00'
            except ValueError:
                tally = '00'
    else:
        tally = '00'
    return tally

def increase_tally(tally_file='../tally.txt'):
    # Read the current tally (as a string)
    tally_str = read_tally(tally_file)
    
    # Convert the tally string to an integer, increment it, then convert it back to a string
    tally = str(int(tally_str) + 1)
    
    # Write the updated tally back to the file
    with open(tally_file, 'w') as file:
        file.write(tally)

    return tally

def decrease_tally(tally_file='../tally.txt'):
    # Read the current tally (as a string)
    tally_str = read_tally(tally_file)
    
    # Convert the tally string to an integer, increment it, then convert it back to a string
    tally = str(int(tally_str) - 1)
    
    # Write the updated tally back to the file
    with open(tally_file, 'w') as file:
        file.write(tally)

    return tally