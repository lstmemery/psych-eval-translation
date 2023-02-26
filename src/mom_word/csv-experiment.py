"""Extract all tables from a Word document and save each as a CSV file."""

import csv
from collections import OrderedDict

import docx
from jinja2 import Template


def convert_table_to_dicts(docx2txt) -> dict[str, OrderedDict[str, dict[str, str | float]]]:
    """Convert a table to an array of rows."""

    table_dicts = {}
    subtest_found = False
    for table in docx2txt.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text == 'Full Scale IQ':
                    table_dicts['composite'] = extract_composite_score_summary(table)
                elif cell.text == 'Pseudoword Decoding' and not subtest_found:
                    subtest_found = True
                    table_dicts['subtest'] = extract_subtest_score_summary(table)
                elif cell.text == 'Oral Discourse Comprehension':
                    table_dicts['component'] = extract_component_score_summary(table)

    return table_dicts


def extract_composite_score_summary(table) -> OrderedDict[str, dict[str, str | float]]:
    """Extract the composite score summary from a table."""
    flattened = flatten_list_of_lists([[cell.text for cell in row.cells] for row in table.rows])
    header_start = flattened.index('Composite')
    header_end = flattened.index('SEM') + 1  # SEM appears twice in the table
    header_length = header_end - header_start + 1

    score = OrderedDict()
    header = flattened[header_start:header_end + 1]
    for row in range(header_end + 1, len(flattened), header_length):
        score[flattened[row]] = {name: value for name, value in zip(header, flattened[row:row + header_length])}
    return score


def extract_subtest_score_summary(table) -> OrderedDict[str, dict[str, str | float]]:
    """Extract the subtest score summary from a table."""
    flattened = flatten_list_of_lists([[cell.text for cell in row.cells] for row in table.rows])
    header_start = flattened.index('Subtest')
    header_end = flattened.index('Growth\nScore') + 1
    header_length = header_end - header_start + 1

    score = OrderedDict()
    header = flattened[header_start:header_end + 1]
    for row in range(header_end + 1, len(flattened), header_length):
        subtest_dict = {name: value for name, value in zip(header, flattened[row:row + header_length])}
        subtest_dict['Qualitative Description'] = get_qualitative_description(float(subtest_dict['Standard\nScore']))
        score[flattened[row]] = subtest_dict
        if subtest_dict['Subtest'] == 'Math Fluency-Multiplication':
            return score  # Avoid footnotes
    return score


def extract_component_score_summary(table) -> OrderedDict[str, dict[str, str | float]]:
    """Extract the component score summary from a table."""
    flattened = flatten_list_of_lists([[cell.text for cell in row.cells] for row in table.rows])
    header_start = flattened.index('Subtest Component')
    header_end = flattened.index('Qualitative\nDescription') + 1
    header_length = header_end - header_start + 1

    score = OrderedDict()
    header = flattened[header_start:header_end + 1]

    allowed_components = {
        'Receptive Vocabulary',
        'Oral Discourse Comprehension',
        'Sentence Combining',
        'Sentence Building',
        'Expressive Vocabulary',
        'Oral Word Fluency',
        'Sentence Repetition',
    }

    for row in range(header_end + 1, len(flattened), header_length):
        if flattened[row]  in allowed_components:
            subtest_dict = {name: value for name, value in zip(header, flattened[row:row + header_length])}
            score[flattened[row]] = subtest_dict
    return score


def get_qualitative_description(score: float) -> str:
    if score > 145:
        return 'Very Superior'
    elif score >= 131:
        return 'Superior'
    elif score >= 116:
        return 'Above Average'
    elif score >= 85:
        return 'Average'
    elif score >= 70:
        return 'Below Average'
    elif score >= 55:
        return 'Low'
    else:
        return 'Very Low'


def flatten_list_of_lists(list_of_lists: list[list[str]]) -> list[str]:
    """Flatten a list of lists into a single list."""
    return [item for sublist in list_of_lists for item in sublist]


if __name__ == '__main__':
    with open('../../data/HL WISC WIAT deidentified.docx', 'rb') as f:
        docx2txt = docx.Document(f)
        table_dicts = convert_table_to_dicts(docx2txt)

        if table_dicts:
            with open('../../data/template.html') as template_file:
                rendered = Template(template_file.read()).render(
                    composite=table_dicts['composite'],
                    subtest=table_dicts['subtest'],
                    component=table_dicts['component'],
                )
                with open('../../data/composite_score_summary.html', 'w') as f:
                    f.write(rendered)


