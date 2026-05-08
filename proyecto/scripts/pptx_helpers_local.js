function imageSizingCrop(_pathOrData, x, y, w, h) {
  return {
    x,
    y,
    w,
    h,
    sizing: { type: 'crop', x, y, w, h },
  };
}

function imageSizingContain(_pathOrData, x, y, w, h) {
  return {
    x,
    y,
    w,
    h,
    sizing: { type: 'contain', x, y, w, h },
  };
}

function safeOuterShadow(color = '000000', opacity = 0.15, angle = 45, blur = 1.5, distance = 1) {
  return { type: 'outer', color, opacity, angle, blur, distance };
}

function _layoutSizeInInches(pptx) {
  const emuPerInch = 914400;
  const widthEmu = Number(pptx?._presLayout?.width) || 0;
  const heightEmu = Number(pptx?._presLayout?.height) || 0;
  return {
    width: widthEmu ? widthEmu / emuPerInch : 13.333,
    height: heightEmu ? heightEmu / emuPerInch : 7.5,
  };
}

function _extractBoxes(slide) {
  const objects = Array.isArray(slide?._slideObjects) ? slide._slideObjects : [];
  return objects
    .map((obj) => {
      const opts = obj?.options || {};
      const x = Number(opts.x);
      const y = Number(opts.y);
      const w = Number(opts.w);
      const h = Number(opts.h);
      if (![x, y, w, h].every(Number.isFinite) || w <= 0 || h <= 0) return null;
      return {
        objectName: opts.objectName || obj?._type || 'element',
        x,
        y,
        w,
        h,
      };
    })
    .filter(Boolean);
}

function _intersects(a, b) {
  const overlapW = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
  const overlapH = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
  if (overlapW <= 0.01 || overlapH <= 0.01) return false;

  const overlapArea = overlapW * overlapH;
  const aArea = a.w * a.h;
  const bArea = b.w * b.h;
  const minArea = Math.min(aArea, bArea);
  if (minArea > 0 && overlapArea / minArea > 0.95) return false;
  return true;
}

function warnIfSlideHasOverlaps(slide) {
  const boxes = _extractBoxes(slide);
  for (let i = 0; i < boxes.length; i += 1) {
    for (let j = i + 1; j < boxes.length; j += 1) {
      if (_intersects(boxes[i], boxes[j])) {
        console.warn(
          `[layout-warning] slide=${slide?._slideNum ?? '?'} overlap="${boxes[i].objectName}"<->"${boxes[j].objectName}"`,
        );
      }
    }
  }
}

function warnIfSlideElementsOutOfBounds(slide, pptx) {
  const boxes = _extractBoxes(slide);
  const layout = _layoutSizeInInches(pptx);
  boxes.forEach((box) => {
    const exceeds = box.x < 0 || box.y < 0 || box.x + box.w > layout.width || box.y + box.h > layout.height;
    if (exceeds) {
      console.warn(
        `[layout-warning] slide=${slide?._slideNum ?? '?'} out-of-bounds="${box.objectName}" box=(${box.x.toFixed(2)},${box.y.toFixed(2)},${box.w.toFixed(2)},${box.h.toFixed(2)}) layout=(${layout.width.toFixed(2)},${layout.height.toFixed(2)})`,
      );
    }
  });
}

module.exports = {
  imageSizingCrop,
  imageSizingContain,
  safeOuterShadow,
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
};
