export function isProcessAlive(
  pid: number,
  signalProcess: (pid: number, signal: 0) => void = process.kill,
): boolean {
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try {
    signalProcess(pid, 0);
    return true;
  } catch (error) {
    return (error as NodeJS.ErrnoException).code !== "ESRCH";
  }
}
