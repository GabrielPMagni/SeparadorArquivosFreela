from extractMethods import build_file_operation1, OperationResult

def test_operation_1_content_finder():
    nf_content = 'Local Nota Fiscal'
    city_content = '2334'

    op = build_file_operation1()

    city_locator = op.fileAttr.city_locator
    nf_locator = op.fileAttr.nf_locator

    content = f"""
    {nf_locator} {nf_content}
    {city_locator} {city_content}


"""

    opResult = op.operation(content, op.fileAttr)

    assert opResult == OperationResult(nf_content=nf_content, nf_city=city_content)
