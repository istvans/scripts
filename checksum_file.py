"""Print the checksum of a file, calculated using common, standard algorithms.

Algos:
- MD5
- SHA256
- SHA512
"""
import hashlib


def calculate_and_print_file_checksum():
    """Calculate and print the checksums for the user specified file."""
    filename = input("Enter the input file name: ")
    filename = filename.replace('"', '')
    with open(filename, "rb") as file:
        entire_file_as_bytes = file.read()

        md5 = hashlib.md5(entire_file_as_bytes).hexdigest()
        print("MD5:", md5)

        sha256 = hashlib.sha256(entire_file_as_bytes).hexdigest()
        print("SHA256:", sha256)

        sha512 = hashlib.sha512(entire_file_as_bytes).hexdigest()
        print("SHA512:", sha512)


if __name__ == "__main__":
    calculate_and_print_file_checksum()
