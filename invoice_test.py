import invoices

def main():
    invoice = Invoice()
    test_set_bank_details(invoice)

def test_set_name(invoice):
    if type(invoice) is not Invoice:
        print('Invalid parameter: Must be an Invoice class')
        return False
    invoice.show_doc()
    print('=================================================================')
    print('Setting name...')
    invoice.set_name('Neville')
    invoice.show_doc()

def test_input_item(invoice):
    if type(invoice) is not Invoice:
        print('Invalid parameter: Must be an Invoice class')
        return False
    invoice.show_doc()
    print('=================================================================')
    print('Inserting 1 item...')
    invoice.insert_item('Reddam House Primary Python',
            '3:30pm - 4:30pm (1 hour) Python class at Reddam House Primary, Term 1, Week 6',
            '17th July, 2018',
            70)
    invoice.show_doc()

def test_set_bank_details(invoice):
    if type(invoice) is not Invoice:
        print('Invalid parameter: Must be an Invoice class')
        return False
    invoice.show_doc()
    print('=================================================================')
    print('Changing bank details...')
    invoice.set_bank_details('nathan nguyen lam', '121212', '87654321')
    invoice.show_doc()
