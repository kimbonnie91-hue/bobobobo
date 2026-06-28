import React from 'react';
import {
  AbsoluteFill,
  interpolate,
  spring,
  useCurrentFrame,
  useVideoConfig,
} from 'remotion';

export const HelloWorld: React.FC<{title: string}> = ({title}) => {
  const frame = useCurrentFrame();
  const {fps} = useVideoConfig();

  const scale = spring({frame, fps, config: {damping: 200}});
  const opacity = interpolate(frame, [0, 30], [0, 1], {
    extrapolateRight: 'clamp',
  });
  const subtitle = interpolate(frame, [30, 60], [0, 1], {
    extrapolateLeft: 'clamp',
    extrapolateRight: 'clamp',
  });

  return (
    <AbsoluteFill
      style={{
        background: 'linear-gradient(135deg,#9070C8,#C080A0)',
        justifyContent: 'center',
        alignItems: 'center',
        fontFamily: 'sans-serif',
      }}
    >
      <div
        style={{
          transform: `scale(${scale})`,
          opacity,
          color: '#fff',
          fontSize: 120,
          fontWeight: 800,
          letterSpacing: '-0.03em',
        }}
      >
        {title}
      </div>
      <div
        style={{
          opacity: subtitle,
          color: '#ffffffcc',
          fontSize: 44,
          marginTop: 24,
        }}
      >
        고객 여정 대시보드
      </div>
    </AbsoluteFill>
  );
};
