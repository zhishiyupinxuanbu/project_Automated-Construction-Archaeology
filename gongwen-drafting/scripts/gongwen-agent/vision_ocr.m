#import <Foundation/Foundation.h>
#import <Vision/Vision.h>
#import <ImageIO/ImageIO.h>
#import <CoreGraphics/CoreGraphics.h>

static NSString *RecognizeText(NSString *imagePath) {
    NSURL *url = [NSURL fileURLWithPath:imagePath];
    CGImageSourceRef source = CGImageSourceCreateWithURL((__bridge CFURLRef)url, NULL);
    if (!source) {
        return @"";
    }
    CGImageRef image = CGImageSourceCreateImageAtIndex(source, 0, NULL);
    CFRelease(source);
    if (!image) {
        return @"";
    }

    VNRecognizeTextRequest *request = [[VNRecognizeTextRequest alloc] initWithCompletionHandler:nil];
    request.recognitionLevel = VNRequestTextRecognitionLevelAccurate;
    request.usesLanguageCorrection = YES;
    if (@available(macOS 11.0, *)) {
        request.recognitionLanguages = @[@"zh-Hans", @"zh-Hant", @"en-US"];
    }

    VNImageRequestHandler *handler = [[VNImageRequestHandler alloc] initWithCGImage:image options:@{}];
    NSError *error = nil;
    BOOL ok = [handler performRequests:@[request] error:&error];
    CGImageRelease(image);
    if (!ok) {
        if (error) {
            fprintf(stderr, "Vision OCR failed for %s: %s\n",
                    [imagePath UTF8String], [[error localizedDescription] UTF8String]);
        }
        return @"";
    }

    NSMutableArray<NSString *> *lines = [NSMutableArray array];
    for (VNRecognizedTextObservation *observation in request.results) {
        NSArray<VNRecognizedText *> *candidates = [observation topCandidates:1];
        if ([candidates count] > 0) {
            NSString *line = candidates[0].string;
            if ([line length] > 0) {
                [lines addObject:line];
            }
        }
    }
    return [lines componentsJoinedByString:@"\n"];
}

int main(int argc, const char * argv[]) {
    @autoreleasepool {
        for (int i = 1; i < argc; i++) {
            NSString *path = [NSString stringWithUTF8String:argv[i]];
            printf("===PAGE:%s===\n", [[path lastPathComponent] UTF8String]);
            NSString *text = RecognizeText(path);
            printf("%s\n", [text UTF8String]);
        }
    }
    return 0;
}
