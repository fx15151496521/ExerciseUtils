package exercise;

import com.eisgroup.exercise.POIUtil;
import com.eisgroup.model.ToolDto;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

/**
 * @Description:
 * @Date: 2019/10/31 14:57
 */
public class POIUtilTest {

    @Test
    public void parseExcel() {
        // POIUtil poiUtil = new POIUtil();
        // POIUtil.parseExcel("POITest-one.xlsx", ToolDto.class);
        List<ToolDto> list = new ArrayList<>();
        ToolDto toolDto = new ToolDto();
        toolDto.setIndex("1");
        toolDto.setColor("红色");
        toolDto.setCountPrice("90");
        toolDto.setDesc("test");
        toolDto.setPartyPrice("80");
        toolDto.setPartyPrice("123");
        toolDto.setType("test");
        toolDto.setPrice("70");
        toolDto.setTypeNumber("45ADFA344");
        list.add(toolDto);
        POIUtil.exportExcel(list, "test.xls", ToolDto.class, "D:/study environment");
    }
}