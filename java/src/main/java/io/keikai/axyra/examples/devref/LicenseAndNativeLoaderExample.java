package io.keikai.axyra.examples.devref;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.License;
import io.keikai.axyra.sheets.LicenseInfo;
import io.keikai.axyra.sheets.Workbook;

import java.nio.file.Path;

public final class LicenseAndNativeLoaderExample implements Example {
    @Override
    public String name() {
        return "dev-ref/license-and-native-loader";
    }

    @Override
    public String description() {
        return "Verify native loading and inspect the active license";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook ignored = Workbook.create()) {
            LicenseInfo info = Workbook.licenseStatus();
            System.out.println("Axyra native library loaded successfully");
            System.out.println("License state: " + info.state());
            System.out.println("PDF enabled: " + License.isFeatureEnabled("pdf_render"));
        }
    }
}
