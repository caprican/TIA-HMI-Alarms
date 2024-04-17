using System.Collections.Generic;

using Siemens.Engineering;
using Siemens.Engineering.AddIn;
using Siemens.Engineering.AddIn.Menu;

namespace TIA_Extract
{
    public class AddInProvider : ProjectTreeAddInProvider
    {
        /// <summary>
        /// The instance of TIA Portal in which the Add-In works.
        /// <para>Enables Add-In to interact with TIA Portal.</para>
        /// </summary>
        private readonly TiaPortal _tiaPortal;

        /// <summary>
        /// The constructor of the FunctionConverterAddInProvider.
        /// <para>- Creates an instance of the class FunctionConverterAddInProvider.</para>
        /// <para>- Called when a right-click is performed in TIA Portal.</para>
        /// </summary>
        /// <param name="tiaPortal">
        /// Represents the instance of TIA Portal in which the Add-In will work.
        /// </param>
        public AddInProvider(TiaPortal tiaPortal)
        {
            _tiaPortal = tiaPortal;
        }

        /// <summary>
        /// The method is provided to include the Add-In
        /// in the context menu of TIA Portal.
        /// </summary>
        /// <typeparam name="ContextMenuAddIn">
        /// The Add-In will be displayed as a new item
        /// in the context menu of TIA Portal.
        /// </typeparam>
        /// <returns>
        /// A new instance of the class FunctionConverterContextMenuAddIn will be created
        /// which implements the main functionality of the Add-In.
        /// </returns>
        protected override IEnumerable<ContextMenuAddIn> GetContextMenuAddIns()
        {
            yield return new SKFAlarmsAddIn(_tiaPortal);
        }

    }
}
