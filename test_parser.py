============================================================
列名                  类型         数字求和              
============================================================
姓名                  text         (无数字)             
年龄                  numeric      156.00              
销售额                numeric      12,345.67           
============================================================from excel_parser import parse_and_calculate, batch_process_xls_files, export_results

def main():
    print("=" * 60)
    print("          XLS文件中文乱码处理工具 - 测试")
    print("=" * 60)
    
    while True:
        print("\n请选择操作:")
        print("1. 处理单个XLS文件")
        print("2. 批量处理文件夹中的XLS文件")
        print("3. 退出")
        
        choice = input("请输入选择 (1/2/3): ")
        
        if choice == '1':
            file_path = input("请输入XLS文件路径: ").strip()
            result = parse_and_calculate(file_path)
            if result:
                export_results([result], 'single_result.txt')
                print("\n处理完成！结果已保存到 single_result.txt")
        
        elif choice == '2':
            folder_path = input("请输入文件夹路径: ").strip()
            results = batch_process_xls_files(folder_path)
            if results:
                export_results(results, 'batch_results.txt')
                print(f"\n处理完成！共处理 {len(results)} 个文件，结果已保存到 batch_results.txt")
        
        elif choice == '3':
            print("退出程序...")
            break
        
        else:
            print("无效选择，请重新输入")

if __name__ == '__main__':
    main()