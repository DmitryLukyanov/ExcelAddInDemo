using System.Drawing;
using System.Windows.Forms;

namespace ExcelAddIn
{
    public partial class SuggestionBox : UserControl
    {
        public SuggestionBox()
        {
            InitializeComponent();
        }

        public override string Text
        {
            get => TextBox.Text;
            set => TextBox.Text = value;
        }

        public override Font Font
        {
            get => TextBox.Font;
            set => TextBox.Font = value;
        }
    }
}
