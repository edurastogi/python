def build_key_index_and_header(csv_path, keycol):
    """Build a mapping from keycol to file offset for a CSV file, and return header."""
    index = {}
    with open(csv_path, 'r', newline='', encoding='utf-8') as f:
        header_line = f.readline()
        header = header_line.strip().split(',')
        key_idx = header.index(keycol)
        while True:
            offset = f.tell()
            line = f.readline()
            if not line:
                break
            row = line.strip().split(',')
            if len(row) > key_idx:
                index[row[key_idx]] = offset
    return index, header

def get_row_dict_by_offset(csv_path, offset, header):
    with open(csv_path, 'r', newline='', encoding='utf-8') as f:
        f.seek(offset)
        line = f.readline()
        values = line.strip().split(',')
        return dict(zip(header, values))

