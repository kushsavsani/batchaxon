setBatchMode(true);
print("--- Fiji Macro Started in Batch Mode ---");

arg = getArgument();
args = split(arg, ",");
inputImagePath = args[0];
outputRoiPath = args[1];
print("  - Input Image: " + inputImagePath);
print("  - Output ROI File: " + outputRoiPath);

print("Opening image and converting overlay...");
open(inputImagePath);
if (nImages() == 0) { exit("FATAL: Image failed to open."); }

run("To ROI Manager");
count = roiManager("count");
print("Found " + count + " ROIs.");

if (count == 0) { exit("FATAL: No outline was found."); }

print("Selecting first ROI and saving...");
roiManager("select", 0);
roiManager("save", outputRoiPath);

if (File.exists(outputRoiPath)) {
    print("SUCCESS: Single .roi file created.");
} else {
    print("FATAL: Failed to save .roi file.");
}
print("--- Fiji Macro Finished ---");