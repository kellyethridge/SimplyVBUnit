** SimplyVBUnit **

version 4.0.2
	* Added support for User Defined Type arrays.
	* Fixed category list. Wasn't reloading selected categories correctly.

version 4.0.1
	* Optimized filling treeview with tests.
	* Tests can optionally be sorted. Currently sorts by name.
	* Can set a custom test comparer for custom sorting.
	* Added error handling during ITestCaseSource.GetTestCases to prevent OCX from crashing project.
	* Optimized progress bar to update only during DoEvents.

version 4.0
	* #Version 4 breaks backward compatibility#
	* Changed ITestCastData to ITestCaseSource. The method signature
	  remains the same. So after changing the method name, the code
	  within the method should continue to work as expected.
	* Removed TestName modifier from the Expect method used when building
	  test cases in the ITestCaseSource implementation.
	* Added categorization capabilities in the framework and component.
	  By implementing the ICategorizable interface, the fixture and tests
	  within the fixture can be categorized.
	* The UI component 'SimplyVBUnit.Component.ocx(SimplyVBComp)' now
	  contains the framework internally, so no need to have a reference
	  to the SimplyVBUnit framework dll when using the Form method of
	  managing and running unit-tests.
	* Added additional constraints.
	* Fixed FixtureTeardown premature calling.
	* The licensing has been changed to MIT licensing.

** Documentation will be added to https://sourceforge.net/p/simplyvbunit/wiki/Home/
   as time progresses. Keep checking back for any updates.
** If you have any questions they can be posted at https://sourceforge.net/p/simplyvbunit/discussion/

** Before running any unit-tests, be sure to set the VB6 IDE Error Trapping to 'Break on Unhandled Errors'.
** This setting can be found in Tools->Options->General Tab.

** Quick Start **
	After installing SimplyVBUnit a project template will be added for VB6. The project name is
	SimplyVBUnit Project. Selecting this as your project will set up a form runner and provide
	references to all things SimplyVBUnit. You should rename the project to reflect the code
	that is to be tested.

	Make sure the SimplyVBUnit Project is set as the start-up project.

	* TESTING A DLL
		With your SimplyVBUnit Project loaded, add a new DLL project to the group or add
		an existing DLL project to the group. From within the SimplyVBUnit Project add a reference
		to the new DLL project to be tested.

	* TESTING AN EXE (if you have the source code)
		With your simplyVBUnit Project loaded, add a new EXE or existing EXE project to the group. 
		A reference cannot be set to the EXE, so a method for exposing the EXE classes and modules to
		the unit-tests is by including the EXE modules in the SimplyVBUnit Project as shared modules. 
		So add the new module to the EXE being tested, then add the same module as an EXISTING module 
		to the SimplyVBUnit project. This way unit-tests can have access to private modules in the EXE.

	* ADDING UNIT-TESTS
		Add classes within the SimplyVBUnit Project that will contain the test code to be run. As classes
		are added to the project, tell SimplyVBUnit about the unit-test class by updating the Form_Load event
		and adding a new instance using the AddTest method:

			Private Sub Form_Load()
				AddTest New MyUnitTests
			End Sub

		When SimplyVBUnit is started, it will discover unit-tests on the classes passed into the AddTest method.