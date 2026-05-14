using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace finalproje2_beyza_ugur_21907047
{
    public partial class Component3 : Component
    {
        public Component3()
        {
            InitializeComponent();
        }

        public Component3(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }
    }
}
