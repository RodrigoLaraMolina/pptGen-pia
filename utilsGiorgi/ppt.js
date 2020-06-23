import pptxgen from "pptxgenjs";

export function createPPT() {
  let pptx = new pptxgen();

  function addSlide(imageUrl, options = {}) {
    var slide = pptx.addSlide();

    slide.addImage({
      data: imageUrl, 
      x: 0, 
      y: 0, 
      w: 10.64, 
      h: 6,
      ...options
    });
  }

  function download(name) {
    pptx.writeFile(name + '.pptx');
  }

  return {
    download,
    addSlide
  }
}