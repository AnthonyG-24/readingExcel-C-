using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Menus
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string fileName = "";

        Excel.Application myApp;
        Excel.Workbook myBook;
        Excel.Worksheet mySheet;
        Excel.Range myRange;

        public MainWindow()
        {
            InitializeComponent();

            Excel.Application myApp = new Excel.Application();
            myApp.Visible = false;

            myApp = new Excel.Application();
        }
        private void Menu_Open(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();

            openDialog.Title = "Select a file to proccess";
            openDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (openDialog.ShowDialog() == true)
            {
                fileName = openDialog.FileName;
                lblBig.Content = "Select data and continue";
            }
        }

        private void Menu_Exit(object sender, RoutedEventArgs e)
        {
            MessageBoxResult answer;
            answer = MessageBox.Show("Do you really want to Exit?", "", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (answer == MessageBoxResult.Yes)
                Application.Current.Shutdown();
        }

        private void fileExist()
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            fileName = openDialog.FileName;
            if (!System.IO.File.Exists(fileName))
            {
                MessageBoxResult result = MessageBox.Show("Error in opening the file " + fileName, "Close Window", MessageBoxButton.OK, MessageBoxImage.Error);
                if (result == MessageBoxResult.OK)
                {
                    Close();
                }

                return;
            }
        }

        private void RunReport(object sender, RoutedEventArgs e)
        { 
            fileExist();

            string optionValue = "";
            string optionItem = "";
            string optionInfo = "";
            string msg = "";

            if ((bool)valueLarge.IsChecked)
            {
                optionValue = "valueLarge";
            } else if ((bool)valueSmall.IsChecked)
            {
                optionValue = "valueSmall";
            } else if ((bool)valueAll.IsChecked)
            {
                optionValue = "valueAll";
            }

            if ((bool)item.IsChecked)
            {
                optionItem = "item";
            } else if ((bool)itemRep.IsChecked)
            {
                optionItem = "itemRep";
            } else if ((bool)itemRegion.IsChecked)
            {
                optionItem = "itemRegion";
            }

            if ((bool)itemSold.IsChecked)
            {
                optionInfo = "itemSold";

            } else if ((bool)itemRev.IsChecked)
            {
                optionInfo = "itemRev";
            }

            //units
            Dictionary<string, int> units_sold = new Dictionary<string, int>();
            Dictionary<string, int> rep_sold = new Dictionary<string, int>();
            Dictionary<string, int> reg_sold = new Dictionary<string, int>();

            //revenue
            Dictionary<string, int> itemRev_sold = new Dictionary<string, int>();
            Dictionary<string, int> salesRepRev_sold = new Dictionary<string, int>();
            Dictionary<string, int> regionRev_sold = new Dictionary<string, int>();

            // Units
            string next_item;
            int next_units;

            string next_rep;
            int next_units2;

            string next_reg;
            int next_units3;

            // Rev
            string next_item2;
            int next_revenue;

            string next_rep2;
            int next_revenue2;

            string next_reg2;
            int next_revenue3;

            myApp = new Excel.Application();
            myApp.Visible = false;
            myBook = myApp.Workbooks.Open(fileName);
            mySheet = myBook.Sheets[1];
            myRange = mySheet.UsedRange;

            int r;
            for (r = 2; r <= myRange.Rows.Count; r++)
            {
                next_item = (string)(mySheet.Cells[r, 4] as Excel.Range).Value;
                next_units = (int)(mySheet.Cells[r, 5] as Excel.Range).Value;

                if (!units_sold.ContainsKey(next_item))
                {
                    units_sold.Add(next_item, next_units);
                } else
                {
                    units_sold[next_item] += next_units;
                }
            }

            int r2;
            for (r2 = 2; r2 <= myRange.Rows.Count; r2++)
            {
                next_rep = (string)(mySheet.Cells[r2, 3] as Excel.Range).Value;
                next_units2 = (int)(mySheet.Cells[r2, 5] as Excel.Range).Value;
            
                if (!rep_sold.ContainsKey(next_rep))
                {
                    rep_sold.Add(next_rep, next_units2);
                } else
                {
                    rep_sold[next_rep] += next_units2;
                }
            }

            int r3;
            for (r3 = 2; r3 <= myRange.Rows.Count; r3++)
            {
                next_reg = (string)(mySheet.Cells[r3, 2] as Excel.Range).Value;
                next_units3 = (int)(mySheet.Cells[r3, 5] as Excel.Range).Value;

                if (!reg_sold.ContainsKey(next_reg))
                {
                    reg_sold.Add(next_reg, next_units3);
                }
                else
                {
                    reg_sold[next_reg] += next_units3;
                }
            }

            int r4;
            for (r4 = 2; r4 <= myRange.Rows.Count; r4++)
            {
                next_item2 = (string)(mySheet.Cells[r4, 4] as Excel.Range).Value;
                next_revenue = (int)(mySheet.Cells[r4, 7] as Excel.Range).Value;

                if (!itemRev_sold.ContainsKey(next_item2))
                {
                   itemRev_sold.Add(next_item2, next_revenue);
                }
                else
                {
                    itemRev_sold[next_item2] += next_revenue;
                }
            }

            int r5;
            for (r5 = 2; r5 <= myRange.Rows.Count; r5++)
            {
                next_rep2 = (string)(mySheet.Cells[r5, 3] as Excel.Range).Value;
                next_revenue2 = (int)(mySheet.Cells[r5, 7] as Excel.Range).Value;

                if (!salesRepRev_sold.ContainsKey(next_rep2))
                {
                    salesRepRev_sold.Add(next_rep2, next_revenue2);
                }
                else
                {
                    salesRepRev_sold[next_rep2] += next_revenue2;
                }
            }

            int r6;
            for (r6 = 2; r6 <= myRange.Rows.Count; r6++)
            {
                next_reg2 = (string)(mySheet.Cells[r6, 2] as Excel.Range).Value;
                next_revenue3 = (int)(mySheet.Cells[r6, 7] as Excel.Range).Value;

                if (!regionRev_sold.ContainsKey(next_reg2))
                {
                    regionRev_sold.Add(next_reg2, next_revenue3);
                }
                else
                {
                    regionRev_sold[next_reg2] += next_revenue3;
                }
            }

            switch (optionValue)
            {
                case "valueLarge":
                    if (optionItem == "item" && optionInfo == "itemSold") 
                    {
                        string mostPopular = "Most popular item overall = ";
                        msg += mostPopular;
                    } else if (optionItem == "itemRep" && optionInfo == "itemSold")
                    {
                        string mostPopular = "Sells rep selling the most units = ";
                        msg += mostPopular;
                    } else if (optionItem == "itemRegion" && optionInfo == "itemSold")
                    {
                        string mostPopular = "Region with the most units = ";
                        msg += mostPopular;
                    }
                    else if (optionInfo == "itemRev" && optionItem == "item")
                    {
                        string mostPopular = "Item with the highest revenue = ";
                        msg += mostPopular;
                    }
                    else if (optionInfo == "itemRev" && optionItem == "itemRep")
                    {
                        string mostPopular = "Sells Rep with the highest revenue = ";
                        msg += mostPopular;
                    }
                    else if (optionInfo == "itemRev" && optionItem == "itemRegion")
                    {
                        string mostPopular = "Region with the highest revenue = ";
                        msg += mostPopular;
                    }
                    break;

                case "valueSmall":
                    if (optionItem == "item" && optionInfo == "itemSold")
                    {
                        string leastPopular = "Least popular item overall = ";
                        msg += leastPopular;
                    }
                    else if (optionItem == "itemRep" && optionInfo == "itemSold")
                    {
                        string leastPopular = "Sells rep selling the least units = ";
                        msg += leastPopular;
                    }
                    else if (optionItem == "itemRegion" && optionInfo == "itemSold")
                    {
                        string leastPopular = "Region with the least sells = ";
                        msg += leastPopular;
                    }
                    if (optionItem == "item" && optionInfo == "itemRev")
                    {
                        string leastPopular = "Item with the least revenue = ";
                        msg += leastPopular;
                    }
                    else if (optionItem == "itemRep" && optionInfo == "itemRev")
                    {
                        string leastPopular = "Sells Rep with the least revenue = ";
                        msg += leastPopular;
                    }
                    else if (optionItem == "itemRegion" && optionInfo == "itemRev")
                    {
                        string leastPopular = "Region with the least revenue = ";
                        msg += leastPopular;
                    }
                    break;

                case "valueAll":
                    if (optionItem == "item" && optionInfo == "itemSold")
                    {
                        string allItems = "Units sold for all items:" + "\n";
                        string dash = "-----------------------" + "\n";
                        msg += allItems + dash;
                    } else if (optionItem == "itemRep" && optionInfo == "itemSold")
                    {
                        string allItems = "All the units all the reps sold:" + "\n";
                        string dash = "-----------------------" + "\n";
                        msg += allItems + dash;
                    } else if (optionItem == "itemRegion" && optionInfo == "itemSold")
                    {
                        string allItems = "All the units all the regions sold:" + "\n";
                        string dash = "-----------------------" + "\n";
                        msg += allItems + dash;
                    } else if (optionItem == "item" && optionInfo == "itemRev")
                    {
                        string allItems = "Revenue for all items:" + "\n";
                        string dash = "-----------------------" + "\n";
                        msg += allItems + dash;

                    } else if (optionItem == "itemRep" && optionInfo == "itemRev")
                    {
                        string allItems = "All the revenue all the reps sold:" + "\n";
                        string dash = "-----------------------" + "\n";
                        msg += allItems + dash;
                    } else if (optionItem == "itemRegion" && optionInfo == "itemRev")
                    {
                        string allItems = "All the revenue of all the regions:" + "\n";
                        string dash = "-----------------------" + "\n";
                        msg += allItems + dash;
                    }
                        break;
            }

            switch (optionItem)
            {
                case "item":
                    if (optionValue == "valueLarge" && optionInfo == "itemSold")
                    {
                        string mostPopularItem = "";
                        int maxUnits = units_sold.Values.Max();

                        foreach (KeyValuePair<string, int> item in units_sold)
                        {
                            if (item.Value == maxUnits)
                            {
                                mostPopularItem = item.Key;
                            }
                        }
                        msg += mostPopularItem;

                    }
                    else if (optionValue == "valueSmall" && optionInfo == "itemSold")
                    {
                        string leastPopularItem = "";
                        int minUnits = units_sold.Values.Min();

                        foreach (KeyValuePair<string, int> item in units_sold)
                        {
                            if (item.Value == minUnits)
                            {
                                leastPopularItem = item.Key;
                            }
                        }
                        msg += leastPopularItem;
                    }
                    else if (optionValue == "valueAll" && optionInfo == "itemSold")
                    {
                        foreach (KeyValuePair<string, int> item in units_sold)
                        {
                            msg += " " + item.Key + " - " + item.Value + "\n";
                        }
                    }

                    //item with most/least revenue
                    if (optionValue == "valueLarge" && optionInfo == "itemRev")
                    {
                        string mostPopularItem = "";
                        int maxUnits = itemRev_sold.Values.Max();

                        foreach (KeyValuePair<string, int> item in itemRev_sold)
                        {
                            if (item.Value == maxUnits)
                            {
                                mostPopularItem = item.Key;
                            }
                        }
                        msg += mostPopularItem;
                    }
                    else if (optionValue == "valueSmall" && optionInfo == "itemRev")
                    {
                        string leastPopularItem = "";
                        int minUnits = itemRev_sold.Values.Min();

                        foreach (KeyValuePair<string, int> item in itemRev_sold)
                        {
                            if (item.Value == minUnits)
                            {
                                leastPopularItem = item.Key;
                            }
                        }
                        msg += leastPopularItem;
                    }
                    else if (optionValue == "valueAll" && optionInfo == "itemRev")
                    {
                        foreach (KeyValuePair<string, int> item in itemRev_sold)
                        {
                            msg += " " + item.Key + " -  $ " + item.Value + "\n";
                        }
                    }
                    break;

                case "itemRep":
                    if (optionValue == "valueLarge" && optionInfo == "itemSold")
                    {
                        int maxUnits = rep_sold.Values.Max();
                        string mostPopular = "";
                        foreach (KeyValuePair<string, int> item in rep_sold)
                        {
                            if (item.Value == maxUnits)
                            {
                                mostPopular = item.Key;
                            }
                        }
                        msg += mostPopular;
                    }
                    else if (optionValue == "valueSmall" && optionInfo == "itemSold")
                    {
                        string leastPopular = "";
                        int minUnits = rep_sold.Values.Min();

                        foreach (KeyValuePair<string, int> item in rep_sold)
                        {
                            if (item.Value == minUnits)
                            {
                                leastPopular = item.Key;
                            }
                        }
                        msg += leastPopular;
                    }
                    else if (optionValue == "valueAll" && optionInfo == "itemSold")
                    {
                        foreach (KeyValuePair<string, int> item in rep_sold)
                        {
                            msg += " " + item.Key + " - " + item.Value + "\n";
                        }
                    }

                    if (optionValue == "valueLarge" && optionInfo == "itemRev")
                    {
                        string mostPopularItem = "";
                        int maxUnits = salesRepRev_sold.Values.Max();

                        foreach (KeyValuePair<string, int> item in salesRepRev_sold)
                        {
                            if (item.Value == maxUnits)
                            {
                                mostPopularItem = item.Key;
                            }
                        }
                        msg += mostPopularItem;
                    }
                    else if (optionValue == "valueSmall" && optionInfo == "itemRev")
                    {
                        string leastPopularItem = "";
                        int minUnits = salesRepRev_sold.Values.Min();

                        foreach (KeyValuePair<string, int> item in salesRepRev_sold)
                        {
                            if (item.Value == minUnits)
                            {
                                leastPopularItem = item.Key;
                            }
                        }
                        msg += leastPopularItem;
                    }
                    else if (optionValue == "valueAll" && optionInfo == "itemRev")
                    {
                        foreach (KeyValuePair<string, int> item in salesRepRev_sold)
                        {
                            msg += " " + item.Key + " -  $" + item.Value + "\n";
                        }
                    }
                    break;

                case "itemRegion":
                    if (optionValue == "valueLarge" && optionInfo == "itemSold")
                    {
                        int maxUnits = reg_sold.Values.Max();
                        string mostPopular = "";
                        foreach (KeyValuePair<string, int> item in reg_sold)
                        {
                            if (item.Value == maxUnits)
                            {
                                mostPopular = item.Key;
                            }
                        }
                        msg += mostPopular;
                    } else if (optionValue == "valueSmall" && optionInfo == "itemSold")
                    {
                        string leastPopular = "";
                        int minUnits = reg_sold.Values.Min();

                        foreach (KeyValuePair<string, int> item in reg_sold)
                        {
                            if (item.Value == minUnits)
                            {
                                leastPopular = item.Key;
                            }
                        }
                        msg += leastPopular;
                    }
                    else if (optionValue == "valueAll" && optionInfo == "itemSold")
                    {
                        foreach (KeyValuePair<string, int> item in reg_sold)
                        {
                            msg += " " + item.Key + " - " + item.Value + "\n";
                        }
                    }

                    if (optionValue == "valueLarge" && optionInfo == "itemRev")
                    {
                        int maxUnits = regionRev_sold.Values.Max();
                        string mostPopular = "";
                        foreach (KeyValuePair<string, int> item in regionRev_sold)
                        {
                            if (item.Value == maxUnits)
                            {
                                mostPopular = item.Key;
                            }
                        }
                        msg += mostPopular;
                    }
                    else if (optionValue == "valueSmall" && optionInfo == "itemRev")
                    {
                        string leastPopular = "";
                        int minUnits = regionRev_sold.Values.Min();

                        foreach (KeyValuePair<string, int> item in regionRev_sold)
                        {
                            if (item.Value == minUnits)
                            {
                                leastPopular = item.Key;
                            }
                        }
                        msg += leastPopular;
                    } else if (optionValue == "valueAll" && optionInfo == "itemRev")
                    {
                        foreach (KeyValuePair<string, int> item in regionRev_sold)
                        {
                            msg += " " + item.Key + " - $" + item.Value + "\n";
                        }
                    }
                        break;
            }

           switch (optionInfo)
           {
                case "itemSold":
                    if (optionItem == "item" && optionValue == "valueLarge")
                    {
                        int maxUnits = units_sold.Values.Max();
                        string convertToString = " (" + maxUnits.ToString() + ")";
                        msg += convertToString;

                    } else if (optionItem == "item" && optionValue == "valueSmall")
                    { 
                        int leastUnits = units_sold.Values.Min();
                        string convertToString = " (" + leastUnits.ToString() + ")";
                        msg += convertToString;
                    }
                    
                    if (optionItem == "itemRep" && optionValue == "valueLarge")
                    {
                        int maxUnits = rep_sold.Values.Max();
                        string convertToString = " (" + maxUnits.ToString() + ")";
                        msg += convertToString;
                    } else if ((optionItem == "itemRep" && optionValue == "valueSmall"))
                    {
                        int leastUnits = rep_sold.Values.Min();
                        string convertToString = " (" + leastUnits.ToString() + ")";
                        msg += convertToString;
                    }

                    if (optionItem == "itemRegion" && optionValue == "valueLarge")
                    {
                        int maxUnits = reg_sold.Values.Max();
                        string convertToString = " (" + maxUnits.ToString() + ")";
                        msg += convertToString;
                    }
                    else if ((optionItem == "itemRegion" && optionValue == "valueSmall"))
                    {
                        int leastUnits = reg_sold.Values.Min();
                        string convertToString = " (" + leastUnits.ToString() + ")";
                        msg += convertToString;
                    }
            
                    break;

               case "itemRev":
                    if (optionItem == "item" && optionValue == "valueLarge")
                    {
                        int maxUnits = itemRev_sold.Values.Max();
                        string convertToString = " (" + "$" + maxUnits.ToString() + ")";
                        msg += convertToString;
                    } else if (optionItem == "item" && optionValue == "valueSmall")
                    {
                        int leastUnits = itemRev_sold.Values.Min();
                        string convertToString = " (" + "$" + leastUnits.ToString() + ")";
                        msg += convertToString;
                    } else if (optionItem == "item" && optionValue == "valueAll")
                    {
                        foreach (KeyValuePair<string, int> item in itemRev_sold)
                        {
                            //msg += " " + item.Key + " - " + "$ " + item.Value + "\n";
                        }
                    }

                    if (optionItem == "itemRep" && optionValue == "valueLarge")
                    {
                        int maxUnits = salesRepRev_sold.Values.Max();
                        string convertToString = " ($" + maxUnits.ToString() + ")";
                        msg += convertToString;
                    }
                    else if ((optionItem == "itemRep" && optionValue == "valueSmall"))
                    {
                        int leastUnits = salesRepRev_sold.Values.Min();
                        string convertToString = " ($" + leastUnits.ToString() + ")";
                        msg += convertToString;
                    } else if (optionItem == "item" && optionValue == "valueAll")
                    {
                        foreach (KeyValuePair<string, int> item in itemRev_sold)
                        {
                           // msg += " " + item.Key + " - " + "$ " + item.Value + "\n";
                        }
                    }

                    if (optionItem == "itemRegion" && optionValue == "valueLarge")
                    {
                        int maxUnits = regionRev_sold.Values.Max();
                        string convertToString = " ($" + maxUnits.ToString() + ")";
                        msg += convertToString;
                    }
                    else if ((optionItem == "itemRegion" && optionValue == "valueSmall"))
                    {
                        int leastUnits = regionRev_sold.Values.Min();
                        string convertToString = " ($" + leastUnits.ToString() + ")";
                        msg += convertToString;
                    } else if (optionItem == "item" && optionValue == "valueAll")
                    {
                        foreach (KeyValuePair<string, int> item in regionRev_sold)
                        {
                           // msg += " " + item.Key + " - " + "$ " + item.Value + "\n";
                        }
                    }
                    break;
           }

            lblBig.Content = msg;
            myBook.Close();
            myApp.Quit();
        }
    }
}
