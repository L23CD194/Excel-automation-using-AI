import org.vosk.Model;
import org.vosk.Recognizer;
import javax.sound.sampled.*;

public class Demo {
    public static void main(String[] args) throws Exception {
        Model model = new Model("C:/Users/chand_ogirl4g/OneDrive/Desktop/keerthi/speech_to_text_project/model");

        try (Recognizer recognizer = new Recognizer(model, 16000)) {
            TargetDataLine microphone = null;

            // Define common audio formats to try
            AudioFormat[] formatsToTry = {
                new AudioFormat(16000, 16, 1, true, false),
                new AudioFormat(44100, 16, 1, true, false),
                new AudioFormat(48000, 16, 1, true, false),
                new AudioFormat(44100, 16, 2, true, false),
                new AudioFormat(48000, 16, 2, true, false)
            };

            // Find and select the buds' microphone
            for (Mixer.Info mixerInfo : AudioSystem.getMixerInfo()) {
                if (mixerInfo.getName().toLowerCase().contains("oneplus nord buds 2")) {
                    Mixer mixer = AudioSystem.getMixer(mixerInfo);
                    for (AudioFormat format : formatsToTry) {
                        DataLine.Info info = new DataLine.Info(TargetDataLine.class, format);
                        if (mixer.isLineSupported(info)) {
                            microphone = (TargetDataLine) mixer.getLine(info);
                            System.out.println("Using microphone: " + mixerInfo.getName() + " with format: " + format.toString());
                            break;
                        }
                    }
                }
                if (microphone != null) break;
            }

            // Fallback to default mic
            if (microphone == null) {
                System.err.println("OnePlus Nord Buds 2 microphone not found. Falling back to default.");
                for (AudioFormat format : formatsToTry) {
                    DataLine.Info info = new DataLine.Info(TargetDataLine.class, format);
                    if (AudioSystem.isLineSupported(info)) {
                        microphone = (TargetDataLine) AudioSystem.getLine(info);
                        System.out.println("Using default microphone with format: " + format.toString());
                        break;
                    }
                }
            }

            if (microphone == null) {
                System.err.println("No supported microphone found.");
                return;
            }

            try {
                microphone.open(microphone.getFormat());
                microphone.start();
                System.out.println("Listening... Speak now!");

                byte[] buffer = new byte[4096];
                int nbytes;

                long silenceStart = System.currentTimeMillis();
                boolean speechDetected = false;

                while (true) {
                    nbytes = microphone.read(buffer, 0, buffer.length);
                    if (nbytes > 0) {
                        if (recognizer.acceptWaveForm(buffer, nbytes)) {
                            speechDetected = true;
                        }
                    }

                    // Auto-stop after 3 seconds of silence once speech has started
                    if (speechDetected && (System.currentTimeMillis() - silenceStart > 3000)) {
                        break;
                    }

                    if (nbytes == 0) {
                        silenceStart = System.currentTimeMillis();
                    }
                }

                // Print only the final result
                System.out.println("Final Text: " + recognizer.getFinalResult());

            } finally {
                if (microphone != null) {
                    microphone.close();
                }
            }
        } finally {
            if (model != null) {
                model.close();
            }
        }
    }
}
