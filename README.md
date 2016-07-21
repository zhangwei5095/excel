# excel
基于poi实现的excel导出工具类
可以通过map指定java实体类与excel文件列名的映射关系，
还可以他通过使用注解来指定文件列名和java字段映射
1 基于map指定映射关系使用示例如下
        List<OrderPo> list = new ArrayList<OrderPo>();
        for(int i=0;i<5000;i++){
            OrderPo po = new OrderPo();
            po.setId(i);
            po.setName("jerry-"+i);
            po.setUid(1000+i);
            po.setCreated(new Date());
            list.add(po);
        }
        LinkedHashMap<String,String> map = new LinkedHashMap<String, String>();
        map.put("序号","id");
        map.put("姓名","name");
        map.put("用户ID","uid");
        map.put("创建时间","created");
        ExcelVo<OrderPo> excelVo = new ExcelVo<OrderPo>(map,list);
        ExcelUtil.createExcelFile("d:/util.xlsx", excelVo, true);
        
2 基于注解实现映射关系
public class OrderPo {
    @Excel(columnName = "主键ID")
    private int id;
    @Excel(columnName = "商品名称")
    private String name;
    @Excel(columnName = "用户的ID")
    private Integer uid;
    @Excel(columnName = "下单时间")
    private Date created;
    private String image;
}
 List<OrderPo> list = new ArrayList<OrderPo>();
        for(int i=0;i<65000;i++){
            OrderPo po = new OrderPo();
            po.setId(i);
            po.setName("jerry-"+i);
            po.setUid(1000+i);
            po.setCreated(new Date());
            list.add(po);
        }
ExcelUtil.createExcelFile("d:/util.xlsx", list);
