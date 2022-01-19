import tools

value = ['姓名', '性别', '联系方式', '家庭住址']
print(len(value))
exit()
while True:
    tools.show_menu()

    action = input('请选择操作功能：')
    print('你选择的功能是：{}'.format(action))

    if action in ['1', '2', '3', '4', '5', '6', '9']:
        if action == '1':
            tools.new_employee()
        elif action == '2':
            tools.show_all()
        elif action == '3':
            tools.search_employee()
        elif action == '4':
            tools.del_employee()
        elif action == '5':
            tools.modify_employee()
        elif action == '6':
            tools.export_to_excel()
        elif action == '9':
            tools.read()
    elif action == '0':
        print('退出')
        break
    else:
        print('输入有误，请再次输入')
