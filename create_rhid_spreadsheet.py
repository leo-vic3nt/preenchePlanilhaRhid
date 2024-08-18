import random
import pandas
import argparse
import os.path as path
import datetime


def generate_cpf():
    cpf = [random.randint(0, 9) for _ in range(9)]

    sum_first_digit = sum((cpf[i] * (10 - i) for i in range(9)))
    first_verification_digit = (sum_first_digit * 10 % 11) % 10
    cpf.append(first_verification_digit)

    sum_second_digit = sum((cpf[i] * (11 - i) for i in range(10)))
    second_verification_digit = (sum_second_digit * 10 % 11) % 10
    cpf.append(second_verification_digit)

    cpf_str = "".join(map(str, cpf))

    return cpf_str


def generate_pis():
    pis = [random.randint(0, 9) for _ in range(10)]

    weights = [3, 2, 9, 8, 7, 6, 5, 4, 3, 2]

    sum_verification_digit = sum(pis[i] * weights[i] for i in range(10))
    verification_digit = 11 - (sum_verification_digit % 11)
    if verification_digit == 10 or verification_digit == 11:
        verification_digit = 0
    pis.append(verification_digit)

    pis_str = "".join(map(str, pis))

    return pis_str


def generate_date_string():
    current_date = datetime.datetime.now()
    date_string = current_date.strftime("%d/%m/%Y")
    return date_string


def validate_file_path(file_path, expected_extension):
    if not path.isfile(file_path):
        raise argparse.ArgumentTypeError(f"{file_path} does not exist.")
    if not file_path.lower().endswith(expected_extension):
        raise argparse.ArgumentTypeError(
            f"{file_path} is not of type {expected_extension}"
        )
    return file_path


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Process a xlsx to csv file for rhid use"
    )
    parser.add_argument(
        "input_xlsx_path",
        type=lambda x: validate_file_path(x, ".xlsx"),
        help="Path to the input XLSX file",
    )
    parser.add_argument(
        "output_csv_path",
        type=lambda x: validate_file_path(x, ".csv"),
        help="Path to the output CSV file",
    )

    return parser.parse_args()


def main():
    args = parse_arguments()

    input_xlsx_path = args.input_xlsx_path
    output_csv_path = args.output_csv_path

    xlsx = pandas.read_excel(input_xlsx_path)
    csv = pandas.read_csv(output_csv_path, sep=";")

    xlsx.columns = xlsx.columns.str.lower().str.strip()

    # Name is required by RHID
    xlsx = xlsx.dropna(subset=["nome"])

    xlsx["cpf"] = xlsx["cpf"].apply(lambda x: generate_cpf() if pandas.isna(x) else x)
    xlsx["pis"] = xlsx["pis"].apply(lambda x: generate_pis() if pandas.isna(x) else x)
    xlsx["data de admissão"] = xlsx["data de admissão"].apply(
        lambda x: (
            generate_date_string()
            if pandas.isna(x)
            else pandas.to_datetime(x).strftime("%d/%m/%Y")
        )
    )

    duplicationFound = False
    duplicated_cpf = xlsx[xlsx.duplicated(subset=["cpf"], keep=False)]

    if not duplicated_cpf.empty:
        print("Duplicated CPF rows:")
        print(duplicated_cpf)
        duplicationFound = True
        print("-------------------------------------")

    duplicated_pis = xlsx[xlsx.duplicated(subset=["pis"], keep=False)]
    if not duplicated_pis.empty:
        print("Duplicated PIS rows:")
        print(duplicated_pis)
        duplicationFound = True
        print("-------------------------------------")

    if duplicationFound:
        print("######################################")
        print("Resolve duplications before proceding")
        print("######################################")
        return 1

    csv.columns = csv.columns.str.lower().str.strip()

    for column in xlsx.columns:
        if column in csv.columns:
            csv[column] = xlsx[column]

    csv.to_csv(output_csv_path, sep=";", index=False)

    print("Data transferred successfully.")


if __name__ == "__main__":
    main()
