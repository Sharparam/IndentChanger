/* IndentChangerAddIn by Sharparam <github.com/Sharparam>
 * 
 * Visual Studio add-in to quickly change the current indent mode (tabs or spaces) via tools menu.
 */

using EnvDTE;
using System.Linq;
using Qreed.VisualStudio;

// Using code from http://qreed.codeplex.com/

namespace IndentChangerAddIn
{
    /// <summary>The object for implementing an add-in.</summary>
    public class Connect : VSAddin
    {
        private const string TopMenuText = "Indent Changer AddIn";

        private const string TabsName = "IndentChangerAddInUseTabs";
        private const string TabsText = "Use tabs for indentation";
        private const string TabsTip = "Changes settings to use tabs when indenting C# code";

        private const string SpacesName = "IndentChangerAddInUseSpaces";
        private const string SpacesText = "Use spaces for indentation";
        private const string SpacesTip = "Changes settings to use spaces when indenting C# code";

        protected override VSMenu SetupMenuItems()
        {
            var rootMenu = new VSMenu(this, TopMenuText);
            var setTabs = rootMenu.AddCommand(TabsName, TabsText, TabsTip);
            setTabs.Execute += (sender, args) => SetMode(true);
            var setSpaces = rootMenu.AddCommand(SpacesName, SpacesText, SpacesTip);
            setSpaces.Execute += (sender, args) => SetMode(false);
            return rootMenu;
        }

        private void SetMode(bool insertTabs)
        {
            ApplicationObject.Properties["TextEditor", "CSharp"].Cast<Property>().First(p => p.Name == "InsertTabs").Value = insertTabs;
        }
    }
}
