using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Automation
{
    /// <summary>
    /// POCO for Main Window Component.
    /// </summary>
    public class MainWindowComponent
    {
        /// <summary>
        /// Gets or sets the Window Title.
        /// </summary>
        [Required]
        public string WindowTitle { get; set; }

        /// <summary>
        /// Gets or sets the Application Name.
        /// </summary>
        [Required]
        public string ApplicationName { get; set; }

        /// <summary>
        /// Gets or sets the Home User Control.
        /// </summary>
        public UserControl HomeUserControl { get; set; } = new UserControlHome();

        /// <summary>
        /// Gets or sets the List of Tools.
        /// </summary>
        [Required]
        public List<ToolComponent> ToolList { get; set; }
    }
}
