using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace finalproje2_beyza_ugur_21907047
{
    public partial class Component2 : Component
    {
        public Component2()
        {
            InitializeComponent();
        }

        public Component2(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }
    }
}
