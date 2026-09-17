import Foundation
import CoreGraphics
import ImageIO
import UniformTypeIdentifiers
let args = CommandLine.arguments
let doc = CGPDFDocument(URL(fileURLWithPath: args[1]) as CFURL)!
let pages = args.count > 3 ? args[3...].compactMap(Int.init) : Array(1...doc.numberOfPages)
for n in pages {
 let page = doc.page(at: n)!
 let box = page.getBoxRect(.mediaBox)
 let scale = 96.0 / 72.0
 let w = Int(box.width * scale), h = Int(box.height * scale)
 let ctx = CGContext(data: nil, width: w, height: h, bitsPerComponent: 8, bytesPerRow: w*4, space: CGColorSpaceCreateDeviceRGB(), bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue)!
 ctx.setFillColor(CGColor(gray: 1, alpha: 1));ctx.fill(CGRect(x:0,y:0,width:w,height:h));ctx.scaleBy(x:scale,y:scale);ctx.drawPDFPage(page)
 let dest = CGImageDestinationCreateWithURL(URL(fileURLWithPath: "\(args[2])-\(n).png") as CFURL, UTType.png.identifier as CFString, 1, nil)!
 CGImageDestinationAddImage(dest, ctx.makeImage()!, nil);CGImageDestinationFinalize(dest)
 print("Page \(n): \(w)x\(h)")
}
