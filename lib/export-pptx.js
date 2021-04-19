const pptxgen = require('pptxgenjs')
module.exports = {
  exportPptx: async (pptxPath, xlsxData) => {
    let pres = new pptxgen();
    pres.defineLayout({ name: 'タイトルスライド', width: 13.3, height: 7.5 });
    pres.layout = 'タイトルスライド'
    let sumary = pres.addSlide();
    sumary.addShape(pres.ShapeType.rect, {
      x: 0.0,
      y: '25%',
      w: '100%',
      h: '40%',
      align: 'center',
      fill: { type: 'solid', color: '2f499c' }
    });
    sumary.addShape(pres.ShapeType.rect, {
      x: 0.0,
      y: '85%',
      w: '100%',
      h: '10%',
      align: 'center',
      fill: { type: 'solid', color: '2f499c' }
    });
    sumary.addText('グループ業績会議', {
      y: '35%',
      w: '100%',
      align: 'center',
      fontSize: 80,
      fontFace: '游ゴシック (Body)',
      color: 'ffffff',
      bold: true,
      isTextBox: true,
    });
    sumary.addText('２０２０年12月　第3四半期', {
      y: '55%',
      w: '100%',
      align: 'center',
      fontSize: 40,
      fontFace: '游ゴシック (Body)',
      color: 'ffffff',
      bold: true,
      isTextBox: true,
    });
    sumary.addText('レクストホールディングス株式会社', {
      y: '90%',
      w: '100%',
      align: 'center',
      fontSize: 40,
      fontFace: '游ゴシック (Body)',
      color: 'ffffff',
      bold: true,
      isTextBox: true,
    });
    pres.writeFile(pptxPath);
  }
}