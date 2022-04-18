package com.example.demo.ThisIsRepository;

import java.util.List;
import java.util.Map;

import org.springframework.context.annotation.ComponentScan;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.CrudRepository;
import org.springframework.stereotype.Repository;

import com.example.demo.ThisIsEntity.TestCase;
@Repository
@ComponentScan
public interface Findrepository extends CrudRepository<TestCase, String> {

	
	@Query(value="select DISTINCT th.ad, th.jcl ,th.para, th.system_operation,th.week,\r\n"
			+ "		tdh.step_name, tdh.jcl_program, tdh.open_mode, tdh.dd, tdh.dsn, tdh.disp, tdh.jcl_library, \r\n"
			+ "        tdh.program_library, tdh.actual_file,tdh.copybook\r\n"
			+ "		FROM rehost.testcase_hpe  th\r\n"
			+ "		inner join rehost.testcase_detail_hpe tdh ON\r\n"
			+ "		th.AD=tdh.AD and \r\n"
			+ "		th.jcl=tdh.jcl\r\n"
			+ "        where th.week='W1' or th.week='W2'",nativeQuery=true)
	List<Map<String, String>> findalldata();

	@Query(value="select * from rehost.testcase_db2_hpe where week='W1' or week='W2'",nativeQuery=true)
	List<Map<String, String>> finddb2data();

	@Query(value="select tih.ad,tih.jcl,tih.step_name,tih.jcl_program,tih.dd,tih.pcb_name,tih.seg_name,tih.opt,tih.week\r\n"
			+ ",(ito.oracle_full_name) as table_name\r\n"
			+ "from rehost.testcase_imsdb_hpe tih\r\n"
			+ "inner join rehost.imsdb_to_oracle ito on\r\n"
			+ "tih.pcb_name=ito.ims_dbd and\r\n"
			+ "tih.seg_name=ito.ims_segment\r\n"
			+ "where tih.week='W1' or tih.week='W2'" ,nativeQuery=true)
	List<Map<String, String>> findimsdbdata();

	
	
	
//	@Query(value="select DISTINCT ad,jcl,listagg(db2_include,',')within group (order by db2_include) as db2_include \r\n"
//			+ ",listagg(ims_get,',')within group (order by ims_get) as ims_get\r\n"
//			+ "from rehost.testcase_db_detail where ad is not null group by ad,jcl",nativeQuery=true)
//	List<Map<String, String>> findgetinclude();

//	@Query(value = "select TESTCASE.SYSTEMTYPE,testcase.testcase,testcase.sprint,testcase.system_operation, TESTCASE.AD, TESTCASE.JCL, TESTCASE_Detail.STEP_NAME,\r\n"
//			+ "TESTCASE_Detail.JCLPROGRAM, TESTCASE_Detail.DD,\r\n"
//			+ "TESTCASE_Detail.DSN, TESTCASE_Detail.OPEN_MODE, TESTCASE_DetaiL.DISP\r\n"
//			+ "from TESTCASE\r\n"
//			+ "inner join TESTCASE_Detail \r\n"
//			+ "on TESTCASE.AD=TESTCASE_Detail.AD and \r\n"
//			+ "TESTCASE.jcl=TESTCASE_Detail.jcl\r\n"
//			+ "where testcase_detail.open_mode is null", nativeQuery = true)
//	public List<Map<String, String>> findalldata();

	
	
	
//	@Query(value = "select TESTCASE.SYSTEMTYPE,testcase.testcase,testcase.sprint,testcase.system_operation, TESTCASE.AD, TESTCASE.JCL, TESTCASE_Detail.STEP_NAME,\r\n"
//			+ "TESTCASE_Detail.JCLPROGRAM, TESTCASE_Detail.DD,\r\n"
//			+ "TESTCASE_Detail.DSN, TESTCASE_Detail.OPEN_MODE, TESTCASE_DetaiL.DISP\r\n"
//			+ "from TESTCASE\r\n"
//			+ "inner join TESTCASE_Detail \r\n"
//			+ "on TESTCASE.AD=TESTCASE_Detail.AD and \r\n"
//			+ "TESTCASE.jcl=TESTCASE_Detail.jcl\r\n"
//			+ "where  (testcase_detail.open_mode='I' or testcase_detail.open_mode='IO')"
//			, nativeQuery = true)
//	public List<Map<String, String>> findinput();
//
//
//
//	@Query(value="select DISTINCT TESTCASE.SYSTEMTYPE,testcase.testcase,testcase.sprint,testcase.system_operation, TESTCASE.AD, TESTCASE.JCL, TESTCASE_Detail.STEP_NAME,\r\n"
//			+ "TESTCASE_Detail.JCLPROGRAM, TESTCASE_Detail.DD,\r\n"
//			+ "TESTCASE_Detail.DSN, TESTCASE_Detail.OPEN_MODE, TESTCASE_DetaiL.DISP\r\n"
//			+ ",CHECKPOINT.PASSFORM,CHECKPOINT.IOCHECKLIST,CHECKPOINT.ALLJCL,CHECKPOINT.SENDFILE\r\n"
//			+ "from TESTCASE\r\n"
//			+ "inner join TESTCASE_Detail \r\n"
//			+ "on TESTCASE.AD=TESTCASE_Detail.AD and \r\n"
//			+ "TESTCASE.jcl=TESTCASE_Detail.jcl\r\n"
//			+ "inner join checkpoint\r\n"
//			+ "on TESTCASE.AD=checkpoint.AD and \r\n"
//			+ "TESTCASE.jcl=checkpoint.jcl \r\n"
//			+ "where (testcase_detail.open_mode='O' or testcase_detail.open_mode='IO')\r\n"
//			+ "and (CHECKPOINT.PASSFORM='Y' or CHECKPOINT.IOCHECKLIST='Y' or CHECKPOINT.ALLJCL='Y' or CHECKPOINT.SENDFILE='Y')", nativeQuery = true)
//	public List<Map<String, String>> findoutput();
//	 
//
//
//	@Query(value="select DISTINCT td.ad, td.jcl, td.step_name, td.jclprogram, td.open_mode, td.dd, td.dsn, td.disp, td.jcl_library, td.program_library , \r\n"
//			+ "t.system_operation\r\n"
//			+ "FROM testcase_detail  td\r\n"
//			+ "inner join TESTCASE t ON\r\n"
//			+ "t.AD=td.AD and \r\n"
//			+ "t.jcl=td.jcl",nativeQuery = true)
//	public List<Map<String, String>> findcatalog();
	
	
	

}
