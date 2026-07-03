import Foundation
import Vision
import ImageIO
import CoreGraphics

func recognizeText(in imagePath: String) -> String {
    let url = URL(fileURLWithPath: imagePath)
    guard let source = CGImageSourceCreateWithURL(url as CFURL, nil),
          let image = CGImageSourceCreateImageAtIndex(source, 0, nil) else {
        return ""
    }

    let request = VNRecognizeTextRequest()
    request.recognitionLevel = .accurate
    request.usesLanguageCorrection = true
    if #available(macOS 11.0, *) {
        request.recognitionLanguages = ["zh-Hans", "zh-Hant", "en-US"]
    }

    let handler = VNImageRequestHandler(cgImage: image, options: [:])
    do {
        try handler.perform([request])
    } catch {
        fputs("Vision OCR failed for \(imagePath): \(error)\n", stderr)
        return ""
    }

    return (request.results ?? [])
        .compactMap { $0.topCandidates(1).first?.string }
        .joined(separator: "\n")
}

let paths = Array(CommandLine.arguments.dropFirst())
for path in paths {
    print("===PAGE:\((path as NSString).lastPathComponent)===")
    print(recognizeText(in: path))
}
