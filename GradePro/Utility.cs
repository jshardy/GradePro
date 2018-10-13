using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GradePro
{
    public static class Utility
    {
        public static IEnumerable<Control> GetAllChildren(this Control root)
        {
            var stack = new Stack<Control>();
            stack.Push(root);

            while(stack.Any())
            {
                var next = stack.Pop();
                foreach(Control child in next.Controls)
                    stack.Push(child);
                yield return next;
            }
        }
    }
}
