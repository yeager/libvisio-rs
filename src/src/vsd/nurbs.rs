//! NURBS curve evaluation using De Boor's algorithm.

/// Evaluate a NURBS curve.
///
/// ctrl_pts: list of (x, y, weight) control points
/// knots: knot vector
/// degree: curve degree (typically 3)
/// num_samples: number of output points
///
/// Returns list of (x, y) points along the curve.
pub fn evaluate_nurbs_curve(
    ctrl_pts: &[(f64, f64, f64)],
    knots: &[f64],
    degree: usize,
    num_samples: usize,
) -> Vec<(f64, f64)> {
    let n = ctrl_pts.len();
    if n < 2 {
        return ctrl_pts.iter().map(|p| (p.0, p.1)).collect();
    }

    // Convert to homogeneous weighted coordinates
    let weighted: Vec<[f64; 3]> = ctrl_pts
        .iter()
        .map(|p| [p.0 * p.2, p.1 * p.2, p.2])
        .collect();

    let p = degree;
    let t_min = if p < knots.len() { knots[p] } else { 0.0 };
    let t_max = if n < knots.len() {
        knots[n]
    } else {
        *knots.last().unwrap_or(&1.0)
    };

    if (t_max - t_min).abs() < 1e-10 {
        return vec![
            (ctrl_pts[0].0, ctrl_pts[0].1),
            (ctrl_pts[n - 1].0, ctrl_pts[n - 1].1),
        ];
    }

    let mut result = Vec::with_capacity(num_samples + 1);
    for step in 0..=num_samples {
        let mut t = t_min + (t_max - t_min) * step as f64 / num_samples as f64;
        t = t.max(t_min).min(t_max - 1e-10);

        // Find knot span
        let mut span = p;
        for j in p..n {
            if j + 1 < knots.len() && knots[j + 1] > t {
                span = j;
                break;
            }
            if j == n - 1 {
                span = n - 1;
            }
        }

        // De Boor recursion
        let mut d: Vec<[f64; 3]> = Vec::with_capacity(p + 1);
        for j in 0..=p {
            let idx = (span as isize - p as isize + j as isize) as usize;
            if idx < n {
                d.push(weighted[idx]);
            } else {
                d.push([0.0, 0.0, 1.0]);
            }
        }

        for r in 1..=p {
            for j in (r..=p).rev() {
                let left = (span as isize - p as isize + j as isize) as usize;
                let ki = left + p + 1 - r;
                let alpha = if ki < knots.len() && left < knots.len() {
                    let denom = knots[ki] - knots[left];
                    if denom.abs() > 1e-10 {
                        (t - knots[left]) / denom
                    } else {
                        0.0
                    }
                } else {
                    0.0
                };
                #[allow(clippy::needless_range_loop)]
                for k in 0..3 {
                    d[j][k] = (1.0 - alpha) * d[j - 1][k] + alpha * d[j][k];
                }
            }
        }

        let w = d[p][2];
        if w.abs() > 1e-10 {
            result.push((d[p][0] / w, d[p][1] / w));
        } else if let Some(&last) = result.last() {
            result.push(last);
        }
    }

    result
}
