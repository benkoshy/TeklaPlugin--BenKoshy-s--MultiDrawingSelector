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
using TSM = Tekla.Structures.Model;
using TSG3D = Tekla.Structures.Geometry3d;
using TSD = Tekla.Structures.Drawing;
using System.Collections;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic.FileIO;
using System.IO;


namespace FlatBarTekla
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Dictionary<string, List<string>> flatBarPartNumberDictionary;
        Dictionary<string, List<string>> PlatePartNumberDictionary;
        Dictionary<string, List<string>> MembersNumberDictionary;

        List<string> selectedItems = new List<string>();
        
        public MainWindow()
        {
            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }           
        }

        private void WriteAndPopulateListBox()
        {
            // create and set dictionary
            WriteDictionaryEntries(ref flatBarPartNumberDictionary, ref PlatePartNumberDictionary);
                        
            // populates the profile list box based on whether the user selects a flat bar dictionary or otherwise
            PopulateProfileList();
        }

        private void PopulateProfileList()
        {
            lstProfiles.Items.Clear();
            lstParts.Items.Clear();            
            
            if (comboDictionaryChoice.SelectedItem != null && comboDictionaryChoice.SelectedItem.ToString() == "System.Windows.Controls.ComboBoxItem: FlatBar")
            {
                foreach (string item in flatBarPartNumberDictionary.Keys)
                {
                    lstProfiles.Items.Add(item);
                }
            }
            else if (comboDictionaryChoice.SelectedItem != null && comboDictionaryChoice.SelectedItem.ToString() == "System.Windows.Controls.ComboBoxItem: Plates")
            {
                foreach (string item in PlatePartNumberDictionary.Keys)
                {
                    lstProfiles.Items.Add(item);
                }
            }
            else if (comboDictionaryChoice.SelectedItem != null && comboDictionaryChoice.SelectedItem.ToString() == "System.Windows.Controls.ComboBoxItem: Members")
            {
                foreach (string item in MembersNumberDictionary.Keys)
                {
                    lstProfiles.Items.Add(item);
                }
            }
        }

        private void lstProfiles_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            string whatWasSelected = (string)lstProfiles.SelectedItem;

            // fills the part number list with values from the appropriate dictionary
            if (comboDictionaryChoice.SelectedItem != null && comboDictionaryChoice.SelectedItem.ToString() == "System.Windows.Controls.ComboBoxItem: FlatBar")
            {
                foreach (string stringValue in flatBarPartNumberDictionary[whatWasSelected])
                {
                    lstParts.Items.Add(stringValue);
                }
            }
            else if (comboDictionaryChoice.SelectedItem != null && comboDictionaryChoice.SelectedItem.ToString() == "System.Windows.Controls.ComboBoxItem: Plates")
            {
                foreach (string stringValue in PlatePartNumberDictionary[whatWasSelected])
                {
                    lstParts.Items.Add(stringValue);
                }
            }
            else if (comboDictionaryChoice.SelectedItem != null && comboDictionaryChoice.SelectedItem.ToString() == "System.Windows.Controls.ComboBoxItem: Members")
            {
                foreach (string stringValue in MembersNumberDictionary[whatWasSelected])
                {
                    lstParts.Items.Add(stringValue);
                }
            }



        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            selectedItems.Clear();

            // once selected then remove from the display
            string whatWasSelected = (string)lstProfiles.SelectedItem;


            // add items into the list - this is the list which is used to select the actual good stuff in Tekla
            foreach (object objectpartno in lstParts.SelectedItems)
            {
                string partNo = (string)objectpartno;
                selectedItems.Add(partNo);
            }


            // select the part numbers in the items to be selected list, in the actual model space.
            SelectModelObjects(selectedItems);

        }

        private void WriteDictionaryEntries(ref Dictionary<string, List<string>> flatBarPartNumberDictionary, ref Dictionary<string, List<string>> PlatePartNumberDictionary)
        {
            // What does this program do?
            // 1. It looks through the modelspace for Flatbars and Plates using a selection filter called the "Single_part_drawing_plate_filter"
            // 2. All flatbars (which are legal) and plates (everything that is not a standard flatbar) is put into a dictionary.
            // 2a. Note: (i) If a multi-part drawing for a particular part exists then it is subtracted from the partMark Dicionaries - as opposed to the
            // dictionaries which have all the actual parts in them.
            // 3. All drawing numbers are obtained.
            // 4. The dictionaries are sorted by order increasining in length
            // 5. We know get a dictionary where the key is the dimension, and the values are: part numbers.
            // 6. The parts which  have those part numbers are selected in the model.

            try
            {
                TSM.Model myModel = new TSM.Model();

                // selects objects according to a filter and puts them in an enumerator
                TSM.ModelObjectEnumerator selectedObjects = myModel.GetModelObjectSelector().GetObjectsByFilterName("Single_part_drawing_plate_filter"); // this drawing filter must collect flatbars and plates too // there is a double located somewhere else for this too.
                TSM.ModelObjectEnumerator selectedObjectsMembers = myModel.GetModelObjectSelector().GetObjectsByFilterName("member_filter"); // this drawing filter must collect flatbars and plates too // there is a double located somewhere else for this too.

                // sets up a dictionary 
                Dictionary<string, List<TSM.Part>> flatBarDictionary = new Dictionary<string, List<TSM.Part>>();
                Dictionary<string, List<TSM.Part>> plateDictionary = new Dictionary<string, List<TSM.Part>>();
                Dictionary<string, List<TSM.Part>> membersDictionary = new Dictionary<string, List<TSM.Part>>();

                // gets Drawing numbers so that only parts with drawing numbers are added to the dictionary
                List<string> drawingNumbers = GetDrawingNumbers();
                SetDictionaries(selectedObjects, ref flatBarDictionary, ref plateDictionary, drawingNumbers);
                // sets the membersDictionary - note the duplication with the above
                SetMemberDictionary(selectedObjectsMembers, ref membersDictionary, drawingNumbers);
                
                // sort dicionaries by increasing length
                SortDicionary(flatBarDictionary);
                SortDicionary(plateDictionary);
                SortDicionary(membersDictionary);

                flatBarPartNumberDictionary = GetPartNumber(flatBarDictionary);
                PlatePartNumberDictionary = GetPartNumber(plateDictionary);
                MembersNumberDictionary = GetPartNumber(membersDictionary);

                // if a single-part drawing already exists in a multi drawing, there is no reason to put it in another multi drawing.
                List<string> allPartMarksExistingInMultiDrawing = GetExistingPartMarksFromMultiDrawings();

                // subtracts existing part marks from dictionary
                if (allPartMarksExistingInMultiDrawing.Count > 0)
                {
                    SubtractExistingPartMarksFromDictionary(flatBarPartNumberDictionary, allPartMarksExistingInMultiDrawing);
                    SubtractExistingPartMarksFromDictionary(PlatePartNumberDictionary, allPartMarksExistingInMultiDrawing);
                    SubtractExistingPartMarksFromDictionary(MembersNumberDictionary, allPartMarksExistingInMultiDrawing);
                }

                // if the dictionary doesn't have any values then remove it from the list
                RemoveDictionaryKeysWhereAllDrawingsAreInMultiSheets(flatBarPartNumberDictionary);
                RemoveDictionaryKeysWhereAllDrawingsAreInMultiSheets(PlatePartNumberDictionary);
                RemoveDictionaryKeysWhereAllDrawingsAreInMultiSheets(MembersNumberDictionary);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private static void RemoveDictionaryKeysWhereAllDrawingsAreInMultiSheets(Dictionary<string, List<string>> dictionary)
        {
            if (dictionary.Keys.Count > 0)
            {
                List<string> itemsToRemove = new List<string>();

                foreach (KeyValuePair<string, List<string>> pair in dictionary)
                {
                    if (pair.Value.Count == 0)
                        itemsToRemove.Add(pair.Key);
                }

                foreach (string item in itemsToRemove)
                {
                    dictionary.Remove(item);
                } 
            }
        }

     


        private void SetMemberDictionary(TSM.ModelObjectEnumerator selectedObjectsMembers, ref Dictionary<string, List<TSM.Part>> membersDictionary, List<string> drawingNumbers)
        {
            Dictionary<string, List<TSM.Part>> thicknessDictionary = new Dictionary<string, List<TSM.Part>>();

            // added this line to speed it up
            selectedObjectsMembers.SelectInstances = false;

            while (selectedObjectsMembers.MoveNext())
            {
                // we are going through all the parts in the model. First we check if a drawing exists for that part. If it does exist, then we:
                // get the profile of the flatbar/plate. But we want just the numbers.

                if (DrawingExists(selectedObjectsMembers.Current, drawingNumbers))
                {

                    // gets the profiles of all parts selected from a particular filter 

                    string profile = "";
                    bool gotProperty = selectedObjectsMembers.Current.GetReportProperty("PROFILE", ref profile);  // for length use "LENGTH"

                    // Profiles will be like "PLT10*100. We want just the numbers. The result string produces a number (string actually) like "10 100"


                    if (profile != "")
                    {
                        AddToDictionary(selectedObjectsMembers, ref membersDictionary, profile);
                    }
                }
            }
        }
        
        private static void SelectModelObjects(List<string> itemsToBeSelected)
        {
            TSM.Model myModel = new TSM.Model();
            TSM.ModelObjectEnumerator selectedObjects = myModel.GetModelObjectSelector().GetAllObjects(); // this drawing filter must collect flatbars and plates too // there is a double located somewhere else for this too.
            selectedObjects.SelectInstances = false;

            TSM.UI.ModelObjectSelector selector = new TSM.UI.ModelObjectSelector();
            ArrayList allSelectedPart = new ArrayList();

            while (selectedObjects.MoveNext())
            {
                TSM.Part part = selectedObjects.Current as TSM.Part;
                if (part != null)
                {
                    string selectedObjectPartMark = part.GetPartMark();
                    //selectedObjectPartMark = Regex.Replace(selectedObjectPartMark, @"[\[\].\/]+", "");

                    foreach (string partMark in itemsToBeSelected)
                    {
                        if (selectedObjectPartMark == partMark)
                        {
                            allSelectedPart.Add(part);
                        }
                    }
                }
            }

            selector.Select(allSelectedPart);
        }

        private static void SubtractExistingPartMarksFromDictionary(Dictionary<string, List<string>> flatBarPartNumberDictionary, List<string> allPartMarksExistingInMultiDrawing)
        {
            foreach (string existingPartMark in allPartMarksExistingInMultiDrawing)
            {
                // now check if the part mark is existing in the list within the dictionary

                foreach (KeyValuePair<string, List<string>> entry in flatBarPartNumberDictionary)
                {
                    for (int i = entry.Value.Count - 1; i >= 0; i--)
                    {
                        if (existingPartMark == entry.Value[i])
                        {
                            entry.Value.RemoveAt(i);
                        }
                    }
                }
            }
        }

        private static List<string> GetExistingPartMarksFromMultiDrawings()
        {
            List<TSD.Drawing> multiDrawings = GetMultiDrawings();       // gets a list of all the multi drawings

            List<string> partMarksExistingInMultiDrawings = new List<string>();

            if (multiDrawings.Count > 0)
            {
                partMarksExistingInMultiDrawings = GetListOfExistingPartMarks(multiDrawings);      // we want to know the part marks for all things already drawn in multi drawings
            }

            return partMarksExistingInMultiDrawings;
        }

        private static List<string> GetListOfExistingPartMarks(List<TSD.Drawing> multiDrawings)
        {
            List<string> existingPartMarks = new List<string>();

            TSM.Model myModel = new TSM.Model();

            foreach (TSD.Drawing drawing in multiDrawings)
            {
                TSD.ContainerView contview = drawing.GetSheet();
                TSD.DrawingObjectEnumerator drwgOjbEnum = contview.GetAllObjects(typeof(TSD.Part));
           

                while (drwgOjbEnum.MoveNext())
                {
                    TSD.Part partDrawing = drwgOjbEnum.Current as TSD.Part;

                    TSM.Part part = myModel.SelectModelObject(partDrawing.ModelIdentifier) as TSM.Part;

                    if (part != null)
                    {
                        string partMark = part.GetPartMark();
                        //partMark = Regex.Replace(partMark, @"[\[\].\/]+", "");


                        if (!existingPartMarks.Contains(partMark))
                        {
                            existingPartMarks.Add(partMark);
                        }

                    }
                }
            }

            return existingPartMarks;
        }

        private static List<TSD.Drawing> GetMultiDrawings()
        {
            TSD.DrawingHandler dh = new TSD.DrawingHandler();

            List<TSD.Drawing> multiDrawings = new List<TSD.Drawing>();

            if (dh.GetConnectionStatus())
            {
                TSD.DrawingEnumerator allDrawings = dh.GetDrawings();
                while (allDrawings.MoveNext())
                {
                    TSD.Drawing currentDrawing = allDrawings.Current as TSD.MultiDrawing;

                    if (currentDrawing != null)
                    {
                        multiDrawings.Add(currentDrawing);
                       // MessageBox.Show("The following multidrawing exists" + currentDrawing.Name);
                    }
                }
            }

            if (multiDrawings.Count == 0)
            {
               // MessageBox.Show("There are no multi drawings existing ATM. Please create them.");
            }

            return multiDrawings;
        }

        private static Dictionary<string, List<string>> GetPartNumber(Dictionary<string, List<TSM.Part>> flatBarDictionary)
        {
            Dictionary<string, List<string>> flatBarPartNumberDictionary = new Dictionary<string, List<string>>();

            foreach (KeyValuePair<string, List<TSM.Part>> entry in flatBarDictionary)
            {
                
               

                // MessageBox.Show(finalKey);
                
                foreach (TSM.Part part in entry.Value)
                {
                    string partMark = part.GetPartMark();
                    //partMark = Regex.Replace(partMark, @"[\[\].\/]+", "");
                                        

                    // checks if the dimension is already in there
                    if (!flatBarPartNumberDictionary.ContainsKey(entry.Key))
                    {
                        // if there is no key (i.e. a plate/flat bar dimension) then we need to instantiate a new list containing the part mark
                        // we then add this list to the dictionary
                        List<string> partMarkList = new List<string>();
                        partMarkList.Add(partMark);
                        flatBarPartNumberDictionary.Add(entry.Key, partMarkList);
                    }
                    else
                    {
                        // check if the partmark is already in the dictionary, if not, then add it
                        if (!flatBarPartNumberDictionary[entry.Key].Contains(partMark))
                        {
                            flatBarPartNumberDictionary[entry.Key].Add(partMark);
                        }
                    }
                }
            }

            Dictionary<string, List<string>> reOrderedAndQtyNotedDictionary = SortDictionaryByNumberOfPartsPerProfileAndEditKeyToReflectThis(ref flatBarPartNumberDictionary);

            return reOrderedAndQtyNotedDictionary;
        }

        private static Dictionary<string, List<string>> SortDictionaryByNumberOfPartsPerProfileAndEditKeyToReflectThis(ref Dictionary<string, List<string>> flatBarPartNumberDictionary)
        {
        
            Dictionary<string, List<string>> reOrderedAndQtyNotedDictionary = new Dictionary<string, List<string>>();

            // creating a new dictionary which has the number of parts noted and added besides the profile string
           
            
            // now order this dictionary by the number of parts contained in the list (located in the keys).

            List<KeyValuePair<string, List<string>>> items = (from pair in flatBarPartNumberDictionary orderby pair.Value.Count descending select pair).ToList();

            foreach(    KeyValuePair<string, List<string>> entry in items)
            {
                int numberOfParts = entry.Value.Count;
                string stringNumberOfParts = string.Format("({0})", numberOfParts);
                string finalKey = entry.Key + " " + stringNumberOfParts;

                reOrderedAndQtyNotedDictionary.Add(finalKey, entry.Value);
            }

            return reOrderedAndQtyNotedDictionary;
        }

        private static void SortDicionary(Dictionary<string, List<TSM.Part>> flatBarDictionary)
        {
            foreach (KeyValuePair<string, List<TSM.Part>> entry in flatBarDictionary)
            {
                entry.Value.Sort(new PartComparer());
            }


            //checks if sorted
            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, List<TSM.Part>> entry in flatBarDictionary)
            {
                sb.Append(Environment.NewLine);
                string headerLineForPartDimension = string.Format("The Lengths for {0} follow", entry.Key);
                sb.Append(headerLineForPartDimension);
                sb.Append(Environment.NewLine);

                foreach (TSM.Part bar in entry.Value)
                {
                    double lengthX = -10;
                    bool gotProperty = bar.GetReportProperty("LENGTH", ref lengthX);
                    sb.Append(string.Join(" ", string.Format("{0:N2} ", lengthX)));
                }

            }

            //string assemblyFolder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            //string filename = Path.Combine(assemblyFolder, "test.txt");
            //File.WriteAllText(filename, sb.ToString());

        }
        
        private static void CreateDrawings_test(Dictionary<string, List<TSM.Part>> flatBarDictionary, string p)
        {
            foreach (KeyValuePair<string, List<TSM.Part>> entry in flatBarDictionary)
            {
                ArrayList samePartSizes = new ArrayList(entry.Value);
                TSM.UI.ModelObjectSelector ms = new TSM.UI.ModelObjectSelector();
                ms.Select(samePartSizes);

            }
        }

        private static void SetDictionaries(TSM.ModelObjectEnumerator selectedObjects, ref Dictionary<string, List<TSM.Part>> flatBarDictionary, ref Dictionary<string, List<TSM.Part>> plateDictionary, List<string> drawingNumbers)
        {
            Dictionary<string, List<TSM.Part>> thicknessDictionary = new Dictionary<string, List<TSM.Part>>();

            // added this line to speed it up
            selectedObjects.SelectInstances = false;

            while (selectedObjects.MoveNext())
            {
                // we are going through all the parts in the model. First we check if a drawing exists for that part. If it does exist, then we:
                // get the profile of the flatbar/plate. But we want just the numbers.

                if (DrawingExists(selectedObjects.Current, drawingNumbers))
                {

                    // gets the profiles of all parts selected from a particular filter 

                    string profile = "";
                    bool gotProperty = selectedObjects.Current.GetReportProperty("PROFILE", ref profile);  // for length use "LENGTH"

                    // Profiles will be like "PLT10*100. We want just the numbers. The result string produces a number (string actually) like "10 100"

                    string resultString = string.Join(" ", Regex.Matches(profile, @"\d+").OfType<Match>().Select(m => m.Value));

                    if (resultString != "")
                    {
                        // converts the "100 10" or "10 100" string into a nominal width x thickness string (with the larger always being first
                        string widthBythickness = GetWidthbyThickness(resultString);

                        if (isFlatBar(profile))
                        {
                            // now check whether the widthbythickness is a standard width if so then add to dictionary:
                            if (IsStandardFlatBar(widthBythickness)) // proceeds only if it's a standard flatbar
                            {
                                // does this thickness exist in the dictionary? is the current item in the enumerator gonna turn out as null? hopefully not

                                AddToDictionary(selectedObjects, ref flatBarDictionary, widthBythickness);
                            }
                            else
                            {
                                AddToDictionary(selectedObjects, ref plateDictionary, widthBythickness);
                            }
                        }
                        else
                        {
                            AddToDictionary(selectedObjects, ref plateDictionary, widthBythickness);
                        }
                    }
                }
            }
        }

        private static List<string> GetDrawingNumbers()
        {

            TSD.DrawingHandler handlder = new TSD.DrawingHandler();
            List<string> drawingNumbers = new List<string>();

            if (handlder.GetConnectionStatus())
            {
                TSD.DrawingEnumerator drawingEnum = handlder.GetDrawings();

                while (drawingEnum.MoveNext())
                {

                    string drawingMark = drawingEnum.Current.Mark;

                    // the full stops are removed from all drawing marks
                    drawingMark = Regex.Replace(drawingMark, @"[\[\].]+", "");

                    if (!drawingNumbers.Contains(drawingMark))
                    {
                        drawingNumbers.Add(drawingMark);
                    }
                }
            }

            if (drawingNumbers.Count == 0)
            {
                MessageBox.Show("Error:  No single part drawings have been created.");
            }
            return drawingNumbers;
        }

        private static bool DrawingExists(TSM.ModelObject modelObject, List<string> drawingNumbers)
        {
            // this sub checks whether the part has a single part drawing existing.

            // TSD.DrawingHandler handlder = new TSD.DrawingHandler();

            TSM.Part part = modelObject as TSM.Part;

            if ((part != null))
            {
               // TSD.DrawingEnumerator drawingEnum = handlder.GetDrawings();
                string partMark = part.GetPartMark();
                partMark = Regex.Replace(partMark, @"[\[\].\/]+", "");
                
                if (drawingNumbers.Contains(partMark)) // make it equal the part mark of the model object
                {
                    return true;
                }
            }

            return false;
        }

        class PartComparer : IComparer<TSM.Part>
        {
            public int Compare(TSM.Part x, TSM.Part y)
            {

                double lengthX = -10;
                bool gotProperty = x.GetReportProperty("LENGTH", ref lengthX);

                double lengthY = -10;
                bool gotPropertY = y.GetReportProperty("LENGTH", ref lengthY);

                if (lengthX == lengthY)
                {
                    return 0;
                }
                else if (lengthX == null || lengthY == null)
                {

                    MessageBox.Show("One of the parts is showing a null value for it's length");
                    return 0;
                }
                else if (lengthX > lengthY)
                {
                    return 1;
                }
                else
                {
                    return -1;
                }

            }

        }

        private static bool IsStandardPlate(string profile)
        {
            string matchStringPlate = @"^(PL)\d+$";  // tests to see whether the profile is a flatbar
            Regex r = new Regex(matchStringPlate);
            Match m = r.Match(profile);

            matchStringPlate = @"^(PLT)\d+$";
            Regex r1 = new Regex(matchStringPlate);
            Match m1 = r1.Match(matchStringPlate);

            if (m.Success || m1.Success)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static bool isFlatBar(string profile)
        {

            string matchStringFlatBar = @"^(PL)\d+(\\*)";  // tests to see whether the profile is a flatbar
            Regex r = new Regex(matchStringFlatBar);
            Match m = r.Match(profile);

            matchStringFlatBar = @"^(PLT)\d+(\\*)";
            Regex r1 = new Regex(matchStringFlatBar);
            Match m1 = r1.Match(matchStringFlatBar);

            matchStringFlatBar = @"^(FL)\d+(\\*)";
            Regex r2 = new Regex(matchStringFlatBar);
            Match m2 = r2.Match(profile);

            matchStringFlatBar = @"^(FPL)\d+(\\*)";
            Regex r3 = new Regex(matchStringFlatBar);
            Match m3 = r3.Match(profile);

            matchStringFlatBar = @"^(PLATE)\d+(\\*)";
            Regex r4 = new Regex(matchStringFlatBar);
            Match m4 = r4.Match(profile);


            if (m.Success || m1.Success || m2.Success || m3.Success || m4.Success)
            {
                return true;
            }


            //MessageBox.Show("the first match returned value is: " + m6.Value.ToString() + " success or fail: " + m6.Success + " and the second matched returned value is " + m7.Value.ToString() + " success or fail: " + m7.Success);

            return false;
        }
        
        private static void AddToDictionary(TSM.ModelObjectEnumerator selectedObjects, ref Dictionary<string, List<TSM.Part>> thicknessDictionary, string widthBythickness)
        {


            if (((selectedObjects.Current as TSM.Part) != null))
            {
                if (!thicknessDictionary.ContainsKey(widthBythickness))
                {
                    // if the thickness does not exist than create a new list add the TSM.Part to the list after casting                           
                    //List<TSM.Part> test = new List<TSM.Part>();
                    //test.Add(selectedObjects.Current as TSM.Part);

                    thicknessDictionary.Add(widthBythickness, new List<TSM.Part>());
                    thicknessDictionary[widthBythickness].Add(selectedObjects.Current as TSM.Part);
                }
                else
                {
                    thicknessDictionary[widthBythickness].Add(selectedObjects.Current as TSM.Part);
                }
            }
        }

        private static string GetWidthbyThickness(string resultString)
        {
            int[] profileArray = resultString.Split(' ').Select(s => Convert.ToInt32(s)).ToArray();

            int thickness = profileArray.Min(); // returns the thickness of the plate
            int width = profileArray.Max();

            string widthBythickness = width + "x" + thickness;
            return widthBythickness;
        }

        private static bool IsStandardFlatBar(string widthBythickness)
        {

            string filePath = GetFilePath();
            bool isStandardFlatBarBool = false;

            TextFieldParser parser = new TextFieldParser(filePath);
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");
            while (!parser.EndOfData && (!isStandardFlatBarBool))
            {
                //Process row
                string[] fields = parser.ReadFields();
                foreach (string field in fields)
                {
                    if (widthBythickness == field)
                    {
                        isStandardFlatBarBool = true;
                        break;
                    }

                }
            }
            parser.Close();


            if (isStandardFlatBarBool == false)
            {
                // MessageBox.Show("Non standard flatbar, sorry: " + widthBythickness);
            }

            return isStandardFlatBarBool;


        }

        private static string GetFilePath()
        {
            // deubbing filename
            //string assemblyFolder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            //string filename = Path.Combine(assemblyFolder, "FlatBarList.CSV");
            //return filename;

            // executing file name
            string assemblyFolder = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            string filename = Path.Combine(assemblyFolder, "FlatBarList.CSV");
            return filename;
        }
        
        // refresh button
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            WriteAndPopulateListBox();
        }

        // ok so the user wishes to change selection then immediately repopulate the lists
        private void comboDictionaryChoice_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PlatePartNumberDictionary != null && flatBarPartNumberDictionary !=null)
            {
                PopulateProfileList();    
            }
            
        }
               
        private void Button_ClearClick_2(object sender, RoutedEventArgs e)
        {
            lstParts.Items.Clear();
        }

        private void lstParts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int numberSelected = lstParts.SelectedItems.Count;
            noSelected.Text = numberSelected.ToString();
        }
    }
}
