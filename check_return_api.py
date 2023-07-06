def check_return_api(data):
    search_string = 'error'
    if search_string in data:
        return 'fail'
    else:
        return 'ok'