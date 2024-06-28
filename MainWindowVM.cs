using Prism.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace ExcelReader
{
    class MainWindowVM
    {
        public DelegateCommand ExtractCommand { get; private set; }
        public DelegateCommand LoadSheetCommand { get; private set; }
        public DelegateCommand GetDoubleRangeCommand { get; private set; }
        public DelegateCommand GetCellCommand { get; private set; }
        public FastExcelReader fxr;
        
        public  MainWindowVM()
            {
            ExtractCommand = new DelegateCommand(Extract);
            LoadSheetCommand = new DelegateCommand(LoadSheet);
            GetCellCommand = new DelegateCommand(GetCell);
            GetDoubleRangeCommand = new DelegateCommand(GetDoubleRange);
           
            string fn = "F:\\csharp\\tryexcel.xlsx";
            fxr = new FastExcelReader(fn);
            sheets = new ObservableCollection<SheetItem>();
            foreach (var s in fxr.sheetnames)
                sheets.Add(new SheetItem { Name = s.Key });
        }
        public void LoadSheet()
        {
            if (SelectedSheetItem == null)
                return;
            fxr.LoadSheet(SelectedSheetItem.Name);
        }
        public void GetDoubleRange()
        {
            if (SelectedSheetItem == null)
                return;
            var st = fxr.GetDoubleRange(SelectedSheetItem.Name, RangeToGet);
        }
        public void GetCell()
        {
            if (SelectedSheetItem == null)
                return;
            var st=fxr.GetCell(SelectedSheetItem.Name, CellToGet);
        }
        private SheetItem _selectedItem;
        public SheetItem SelectedSheetItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value;
                
                OnPropertyChanged(nameof(SelectedSheetItem));
                // You can add additional logic here if needed
            }
        }
        private string _celltoget="A0";
        public string CellToGet
        {
            get { return _celltoget; }
            set
            {
                _celltoget = value;

                OnPropertyChanged(nameof(CellToGet));
                // You can add additional logic here if needed
            }
        }
        private string _rangetoget = "D2:D8";
        public string RangeToGet
        {
            get { return _rangetoget; }
            set
            {
                _rangetoget = value;

                OnPropertyChanged(nameof(RangeToGet));
                // You can add additional logic here if needed
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        public ObservableCollection<SheetItem> sheets { get; set; }
        public void Extract()
        {
            fxr.ExtractDoc();
        }
       
    }
    public class SheetItem
    {
        public string Name { get; set; }
    }
  

}  

