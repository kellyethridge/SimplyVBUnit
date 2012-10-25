** SimplyVBUnit **

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