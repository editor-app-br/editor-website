import * as React from "react";
import type { SVGProps } from "react";

/** Brand mark: isometric cube (vector), not the inherited ZIZIYI “O”. */
const SvgLogo = (props: SVGProps<SVGSVGElement>) => (
  <svg
    xmlns="http://www.w3.org/2000/svg"
    fill="none"
    viewBox="0 0 512 512"
    {...props}
  >
    <image href="/editor_cube_transparent.svg" width={512} height={512} />
  </svg>
);

export default SvgLogo;
