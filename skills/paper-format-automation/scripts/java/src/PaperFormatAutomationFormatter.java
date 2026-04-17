import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Path;
import java.nio.file.Paths;

public class PaperFormatAutomationFormatter {
    public static void main(String[] args) throws IOException, InterruptedException, URISyntaxException {
        String input = null;
        String rules = null;
        String output = null;
        for (int i = 0; i < args.length - 1; i++) {
            if ("--input".equals(args[i])) input = args[++i];
            else if ("--rules".equals(args[i])) rules = args[++i];
            else if ("--output".equals(args[i])) output = args[++i];
        }
        if (input == null || rules == null || output == null) {
            throw new IllegalArgumentException("Expected --input <docx> --rules <json> --output <docx>");
        }

        Path jarDir = Paths.get(PaperFormatAutomationFormatter.class.getProtectionDomain().getCodeSource().getLocation().toURI()).getParent();
        Path script = jarDir.getParent().resolve("format_manuscript.py").normalize();

        ProcessBuilder pb = new ProcessBuilder(
            "python",
            script.toString(),
            "--input", input,
            "--rules", rules,
            "--output", output
        );
        pb.inheritIO();
        Process process = pb.start();
        int code = process.waitFor();
        if (code != 0) {
            throw new IOException("Python formatter failed with exit code " + code);
        }
        System.out.println("Formatter launcher completed using Python formatting core.");
    }
}
