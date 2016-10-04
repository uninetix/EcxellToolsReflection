package pl.ciszemar.main;

import java.io.File;
import java.util.ArrayList;

import pl.ciszemar.demo.Employee;
import pl.ciszemar.tools.ExcellTool;

public class ExcellReflectionTool {

	public static void main(String[] args) {

		Employee emp1 = new Employee();
		emp1.setFirstName("Jan");
		emp1.setLastName("Kowalski");
		emp1.setSalary(2345L);
		Employee emp2 = new Employee();
		emp2.setFirstName("Krzysztof");
		emp2.setLastName("Nowak");
		emp2.setSalary(5432L);
		File f = new File("ExcellToolsDemo.xls");
		ArrayList<Employee> empList = new ArrayList<Employee>();
		empList.add(emp1);
		empList.add(emp2);
		ExcellTool excellTool = new ExcellTool();
		
		excellTool.writeObjectList(f, empList);
		

	}

}
