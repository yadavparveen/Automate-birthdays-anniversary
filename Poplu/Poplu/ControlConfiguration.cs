using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Poplu
{
    public class ControlConfiguration
    {
        public string LabelText { get; set; }
        public int LabelXLocation { get; set; }
        public int LabelYLocation { get; set; }
        public Color LabelBackgroundColor { get; set; }
        public Color LabelForColor { get; set; }
        public int GridXLocation { get; set; }
        public int GridYLocation { get; set; }
        public Color GridBackgroundColor { get; set; }
        public Color GridForColor { get; set; }

    }
}
