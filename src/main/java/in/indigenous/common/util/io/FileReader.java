package in.indigenous.common.util.io;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public interface FileReader {

	Map<Object, List<Object>> read(String dir, String fileName) throws IOException;
}
