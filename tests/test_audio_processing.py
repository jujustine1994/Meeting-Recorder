import sys
import os
import struct
import array
import unittest

import unittest.mock
# Mock native audio/encoding libs so tests run without hardware dependencies
unittest.mock.patch.dict('sys.modules', {
    'pyaudiowpatch': unittest.mock.MagicMock(),
    'lameenc': unittest.mock.MagicMock(),
}).start()

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from main import MeetingRecorderApp, _compute_rms


def make_pcm(value: int, n_samples: int = 512) -> bytes:
    """產生 n_samples 個相同振幅的 PCM Int16 bytes"""
    return struct.pack(f"{n_samples}h", *([value] * n_samples))


class TestComputeEqualizeGain(unittest.TestCase):

    def test_mic_quieter_gets_boosted(self):
        sys_frames = [make_pcm(1000)] * 10
        mic_frames = [make_pcm(500)] * 10  # ratio = 2.0, inside cap
        sys_gain, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            sys_frames, mic_frames, filter_silence=False)
        self.assertEqual(sys_gain, 1.0)
        self.assertAlmostEqual(mic_gain, 2.0, places=1)

    def test_sys_quieter_gets_boosted(self):
        sys_frames = [make_pcm(250)] * 10
        mic_frames = [make_pcm(1000)] * 10
        sys_gain, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            sys_frames, mic_frames, filter_silence=False)
        self.assertAlmostEqual(sys_gain, 4.0, places=1)
        self.assertEqual(mic_gain, 1.0)

    def test_gain_capped_at_gain_cap(self):
        sys_frames = [make_pcm(10000)] * 10
        mic_frames = [make_pcm(100)] * 10   # 比值 = 100，超過 cap
        _, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            sys_frames, mic_frames, filter_silence=False, gain_cap=4.0)
        self.assertEqual(mic_gain, 4.0)

    def test_both_silent_returns_no_gain(self):
        sys_frames = [make_pcm(0)] * 10
        mic_frames = [make_pcm(0)] * 10
        sys_gain, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            sys_frames, mic_frames, filter_silence=False)
        self.assertEqual(sys_gain, 1.0)
        self.assertEqual(mic_gain, 1.0)

    def test_one_track_silent_returns_no_gain(self):
        sys_frames = [make_pcm(0)] * 10
        mic_frames = [make_pcm(1000)] * 10
        sys_gain, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            sys_frames, mic_frames, filter_silence=False)
        self.assertEqual(sys_gain, 1.0)
        self.assertEqual(mic_gain, 1.0)

    def test_equal_rms_returns_no_gain(self):
        frames = [make_pcm(500)] * 10
        sys_gain, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            frames, frames, filter_silence=False)
        self.assertEqual(sys_gain, 1.0)
        self.assertEqual(mic_gain, 1.0)

    def test_filter_silence_excludes_quiet_chunks(self):
        # sys: 全部 loud（RMS≈1000）
        # mic: 一半 silent（RMS≈10, 低於閾值100）、一半 loud（RMS≈1000）
        # filter_silence=True 後，mic active_rms ≈ 1000 → 與 sys 等響 → gain ≈ 1.0
        sys_frames = [make_pcm(1000)] * 10
        mic_frames = [make_pcm(10)] * 5 + [make_pcm(1000)] * 5
        _, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            sys_frames, mic_frames, filter_silence=True, gain_cap=10.0)
        self.assertAlmostEqual(mic_gain, 1.0, places=1)

    def test_filter_silence_false_includes_quiet_chunks(self):
        # 不過濾靜音時，平均 RMS 被靜音段壓低，gain > 1
        sys_frames = [make_pcm(1000)] * 10
        mic_frames = [make_pcm(10)] * 5 + [make_pcm(1000)] * 5
        _, mic_gain = MeetingRecorderApp._compute_equalize_gain(
            sys_frames, mic_frames, filter_silence=False, gain_cap=10.0)
        self.assertGreater(mic_gain, 1.5)


class TestApplyGainToPcm(unittest.TestCase):

    def test_gain_1_unchanged(self):
        data = make_pcm(1000)
        self.assertEqual(MeetingRecorderApp._apply_gain_to_pcm(data, 1.0), data)

    def test_gain_doubles_amplitude(self):
        data = make_pcm(1000, n_samples=4)
        result = MeetingRecorderApp._apply_gain_to_pcm(data, 2.0)
        arr = array.array('h')
        arr.frombytes(result)
        self.assertEqual(list(arr), [2000, 2000, 2000, 2000])

    def test_positive_clamp_prevents_overflow(self):
        data = make_pcm(20000, n_samples=4)
        result = MeetingRecorderApp._apply_gain_to_pcm(data, 2.0)
        arr = array.array('h')
        arr.frombytes(result)
        self.assertTrue(all(s == 32767 for s in arr))

    def test_negative_clamp_prevents_overflow(self):
        data = struct.pack("4h", -20000, -20000, -20000, -20000)
        result = MeetingRecorderApp._apply_gain_to_pcm(data, 2.0)
        arr = array.array('h')
        arr.frombytes(result)
        self.assertTrue(all(s == -32768 for s in arr))

    def test_zero_gain_returns_silence(self):
        data = make_pcm(1000, n_samples=4)
        result = MeetingRecorderApp._apply_gain_to_pcm(data, 0.0)
        arr = array.array('h')
        arr.frombytes(result)
        self.assertEqual(list(arr), [0, 0, 0, 0])


if __name__ == "__main__":
    unittest.main()
