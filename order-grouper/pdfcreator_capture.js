/* global ActiveXObject, WScript */

(function () {
    "use strict";

    var argumentsList = WScript.Arguments;
    if (argumentsList.length < 2) {
        WScript.Echo(
            "Usage: pdfcreator_capture.js <destination.pdf> <timeout-seconds> [profile]"
        );
        WScript.Quit(2);
    }

    var destination = argumentsList.Item(0);
    var timeoutSeconds = parseInt(argumentsList.Item(1), 10);
    var profile = argumentsList.length >= 3 ? argumentsList.Item(2) : "";
    var queue = null;

    try {
        queue = new ActiveXObject("PDFCreator.JobQueue");
        queue.Initialize();

        if (!queue.WaitForJob(timeoutSeconds)) {
            WScript.Quit(10);
        }

        var job = queue.NextJob;
        if (profile) {
            job.SetProfileByGuidOrName(profile);
        }
        job.SetProfileSetting("OutputFormat", "Pdf");
        job.SetProfileSetting("ShowProgress", "false");
        job.SetProfileSetting("ShowQuickActions", "false");
        job.SetProfileSetting("ShowAllNotifications", "false");
        job.ConvertTo(destination);

        if (!job.IsFinished || !job.IsSuccessful) {
            throw new Error("PDFCreator conversion was not successful.");
        }

        WScript.Echo(destination);
        WScript.Quit(0);
    } catch (error) {
        WScript.Echo(error.message || String(error));
        WScript.Quit(1);
    } finally {
        if (queue !== null) {
            try {
                queue.ReleaseCom();
            } catch (releaseError) {
                // The process exit will release the COM object if this fails.
            }
        }
    }
}());
