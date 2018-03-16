import javax.script.*;
public class EvalFile {
    public static void main(String[] args) throws Exception {
        // create a script engine manager
        ScriptEngineManager factory = new ScriptEngineManager();
        // create JavaScript engine
        ScriptEngine engine = factory.getEngineByName("JavaScript");
        // evaluate JavaScript code from given file - specified by first argument
        engine.eval(new java.io.FileReader(args[0]));
    }
}
