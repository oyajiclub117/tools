import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;
 
public class ScriptSample1 {
    public ScriptSample1() {
        ScriptEngineManager manager = new ScriptEngineManager();
        ScriptEngine engine = manager.getEngineByName("js");
        
        String script = "print('Hello, World!');";
 
        if (engine != null) {
            try {
                engine.eval(script);
            } catch (ScriptException ex) {
                ex.printStackTrace();
            }
        }
    }
 
    public static void main(String[] args) {
        new ScriptSample1();
    }
}
