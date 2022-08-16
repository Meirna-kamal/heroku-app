from flask_restful import Api, reqparse, abort
from flask import Flask

from csv import writer
import pandas as pd

app = Flask(__name__)
api = Api(app)

novel_args = reqparse.RequestParser()

novel_args.add_argument(
    "الروايه", type=str, help="Name of the novel", required=True)
novel_args.add_argument(
    "المؤلف", type=str, help="Name of the author", required=True)
novel_args.add_argument(
    "البلد", type=str, help="Country of author", required=True)

def convert_excel_to_df():
    df = pd.read_excel('Final_without_links.xlsx')
    del df['Unnamed: 0']
    return df


def format_excel(writer, df):
    df.to_excel(writer, sheet_name='Sheet1')

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    # Implicit format.

    # Add the cell formats.
    format_right_to_left = workbook.add_format({'reading_order': 2})

    # Change the direction for the worksheet.
    worksheet.right_to_left()

    # Make the column wider for visibility and add the reading order format.
    worksheet.set_column('B:D', 20, format_right_to_left)
    writer.save()


def abort_if_index_out_of_range(index, df):
    if index > len(df):
        abort(404, message="Novel Index out of range")


@app.route("/")
def index():
    return ("Best-Arabic-Novels Api")


@app.route('/Novel/<int:novel_id>', methods=['GET'])
def get_novel(novel_id):
    df = convert_excel_to_df()
    abort_if_index_out_of_range(novel_id, df)

    requested_novel_name = f"{ (df.get('الروايه')[novel_id-1])}"
    requested_novel_author = f"{ (df.get('المؤلف')[novel_id-1])}"
    requested_novel_country = f"{ (df.get('البلد')[novel_id-1])}"
    return { ("الروايه"): requested_novel_name,
             ("المؤلف"): requested_novel_author,
             ("البلد"): requested_novel_country}


@app.route('/Novel', methods=['POST'])
def post_novel():
    args = novel_args.parse_args()

    df = convert_excel_to_df()
    df.loc[len(df), df.columns] = args['الروايه'], args['المؤلف'], args['البلد']
    df.index = df.index + 1  # shifting index
    writer = pd.ExcelWriter(
        'Final_without_links.xlsx', engine='xlsxwriter')
    format_excel(writer, df)

    return { ("الروايه"): ( args['الروايه']),
             ("المؤلف"):  (args['المؤلف']),
             ("البلد"):  (args['البلد'])}, 201


@app.route('/Novel/<int:novel_id>', methods=['PUT'])
def put_novel(novel_id):
    args = novel_args.parse_args()
    df = convert_excel_to_df()
    abort_if_index_out_of_range(novel_id, df)

    df.loc[novel_id-1] = args['الروايه'], args['المؤلف'], args['البلد']
    df.index = df.index + 1  # shifting index
    writer = pd.ExcelWriter(
        'Final_without_links.xlsx', engine='xlsxwriter')
    format_excel(writer, df)

    return { ("الروايه"):  (args['الروايه']),
             ("المؤلف"):  (args['المؤلف']),
             ("البلد"):  (args['البلد'])}, 201


@app.route('/Novel/<int:novel_id>', methods=['DELETE'])
def delete_novel(novel_id):
    df = convert_excel_to_df()
    abort_if_index_out_of_range(novel_id, df)

    df = df.drop(df.index[novel_id-1])
    df = df.reset_index(drop=True)
    df.index = df.index + 1
    writer = pd.ExcelWriter(
        'Final_without_links.xlsx', engine='xlsxwriter')
    format_excel(writer, df)
    return '', 204

if __name__ == "__main__":
    app.run()
