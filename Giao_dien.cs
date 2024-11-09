using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Data_progress;
using Validator;
using OfficeOpenXml;
using Microsoft.SqlServer.Server;
namespace QLNV_FTI
{
    public class Giao_dien
    {
        static void Main (string[] args)
        {
            employee_management Emp_Mana = new employee_management();
            string defaultFilePath = @"C:\Users\LENOVO\Desktop\QLNV-FTI.xlsx";

            while (true)
            {
                Console.WriteLine("_____MENU_____");
                Console.WriteLine("Chon cac tinh nang sau: ");
                Console.WriteLine("1.Nhap thong tin nhan vien");
                Console.WriteLine("2.Xuat du lieu ra file excel");
                Console.WriteLine("3.Tim kiem nhan vien theo ID");
                Console.WriteLine("4.Thoat chuong trinh");
                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        Console.Write("Ban muon them bao nhieu nhan vien ? ");
                        if (!int.TryParse(Console.ReadLine(), out int quantity) || quantity < 1 || quantity > 4)
                        {
                            Console.WriteLine("So luong khong hop le!");
                            break;
                        }

                        for (int i = 0; i < quantity; i++)
                        {
                            Console.WriteLine($"Nhap thong tin cho nhan vien thu {i + 1}:");
                            
                            Console.Write("Nhap ten: ");
                            string name = Console.ReadLine();
                            
                            Console.Write("Nhap ID: ");
                            string id = Console.ReadLine();

                            Console.Write("Nhap gioi tinh (Nam/Nu): ");
                            string gender = Console.ReadLine();

                            Console.Write("Nhap tuoi: ");
                            string age = Console.ReadLine();

                            Console.Write("Nhap luong co ban: ");
                            string base_salary = Console.ReadLine();

                            Console.Write("Nhap he so luong: ");
                            string salary_coefficient = Console.ReadLine();

                            Console.Write("Nhap phu cap: ");
                            string allowance = Console.ReadLine();
                            
                            Console.Write("Nhap ma cong doan: ");
                            string pro_id = Console.ReadLine();

                            Console.Write("Nhap ten cong doan: ");
                            string pro_name = Console.ReadLine();

                            Console.Write("Nhap so luong: ");
                            string qty = Console.ReadLine();

                            Console.Write("Nhap gia: ");
                            string price = Console.ReadLine();

                            string thong_tin_nv = Emp_Mana.Employee_Insert(name, id, age, gender, base_salary, salary_coefficient, allowance, pro_id, pro_name, qty , price );
                            Console.WriteLine(thong_tin_nv);

                            
                        }
                        break;

                    case "2":
                        string exportResult = Emp_Mana.ExportToExcel(defaultFilePath);
                        Console.WriteLine(exportResult);
                        break;

                    case "3":
                        Console.Write("Nhap ID cua nhan vien can tim: "); 
                        string SearchID = Console.ReadLine();
                        var EmployeebyID = Emp_Mana.Findemployee(SearchID);
                        if (EmployeebyID.Count > 0 && EmployeebyID.Count <= 50)
                        {
                            Console.WriteLine($"Danh sach nhan vien co ID '{SearchID}':");
                            foreach (var employee in EmployeebyID)
                            {
                                Console.WriteLine($"{employee.ID} - {employee.Name} - {employee.Gender}");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Khong tim thay nhan vien nao !.");
                        }
                        break;

                    case "4":
                        Console.WriteLine("Thoat chuong trinh");
                        return;

                    default:
                        Console.WriteLine("Tuy chon khong hop le. Vui long chon lai");
                        break;


                }           
            }
        }
    }
}
