#version 440

layout(location = 0) in  vec2 qt_TexCoord0;
layout(location = 0) out vec4 fragColor;

layout(std140, binding = 0) uniform buf {
    mat4  qt_Matrix;
    float qt_Opacity;
    float progress;       // 0.0 → 1.0 driven by QML NumberAnimation
} ubuf;

layout(binding = 1) uniform sampler2D source;

#define PI 3.14159265358979

void main() {
    vec2  uv     = qt_TexCoord0;
    vec2  center = vec2(0.5, 0.5);
    vec2  delta  = uv - center;
    float dist   = length(delta);
    float angle  = atan(delta.y, delta.x);

    float p = ubuf.progress;

    // Twist amount: strongest near center, grows with progress
    float twist = p * 5.5 * PI * (1.0 - smoothstep(0.0, 0.65, dist));

    // Contracted distance: UI shrinks to center point
    float contracted = dist * max(0.0, 1.0 - p * 0.99);

    // Reconstruct warped UV
    float twistedAngle = angle + twist;
    vec2  srcUV = center + contracted * vec2(cos(twistedAngle), sin(twistedAngle));
    srcUV = clamp(srcUV, 0.001, 0.999);

    vec4  col  = texture(source, srcUV);

    // Colour shift toward deep blue as vortex deepens
    vec3  tint = mix(col.rgb, vec3(0.05, 0.18, 0.44), p * 0.55);

    // Combined fade: image goes dark quickly past 75 % progress
    float fade = max(0.0, 1.0 - pow(max(0.0, p - 0.0) / 1.0, 1.6));

    fragColor = vec4(tint * fade, col.a * fade) * ubuf.qt_Opacity;
}
