function imageSizingCrop(_pathOrData, x, y, w, h) {
  return { x, y, w, h };
}

function imageSizingContain(_pathOrData, x, y, w, h) {
  return { x, y, w, h };
}

function safeOuterShadow(color = '000000', opacity = 0.15, angle = 45, blur = 1.5, distance = 1) {
  return { type: 'outer', color, opacity, angle, blur, distance };
}

function warnIfSlideHasOverlaps() {}
function warnIfSlideElementsOutOfBounds() {}

module.exports = {
  imageSizingCrop,
  imageSizingContain,
  safeOuterShadow,
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
};
