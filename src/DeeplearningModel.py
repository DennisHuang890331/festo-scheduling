import tensorflow as tf
class CustomSchedule(tf.keras.optimizers.schedules.LearningRateSchedule):
    def __init__(self, d_model, warmup_steps=100):
        super(CustomSchedule, self).__init__()
        self.d_model = d_model
        self.d_model = tf.cast(self.d_model, tf.float32).numpy()
        self.warmup_steps = warmup_steps
    def __call__(self, step):
        arg1 = tf.math.rsqrt(step)
        arg2 = step * (self.warmup_steps ** -1.25)
        return tf.math.rsqrt(self.d_model) * tf.math.minimum(arg1, arg2)
    def get_config(self):
        config = {
        'd_model': self.d_model,
        'warmup_steps': self.warmup_steps,
        }
        return config
class cooling_model:
    def __init__(self):
        try:
            print("類神經網路冷卻模型讀取中......")
            self.model = tf.keras.models.load_model("Cooling_Model.h5")
            print("冷卻模型讀取完畢......")
            print("冷卻模型Summary......")
            # self.model.summary()
            print("\n")
        except(Exception):
            print(Exception)
class heating_model:
    def __init__(self):
        try:
            print("類神經網路加熱模型讀取中......")
            self.model = tf.keras.models.load_model("Heating_Model.h5")
            print("加熱模型讀取完畢......")
            print("加熱模型Summary......")
            # self.model.summary()
            print("\n")
        except(Exception):
            print(Exception)
class LSTM_model:
    def __init__(self):
        try:
            print("loading LSTM temperture predict model ......")
            with tf.keras.utils.CustomObjectScope({'CustomSchedule': CustomSchedule}):
                self.model = tf.keras.models.load_model('DL_LSTM_v1.h5')
            print("Finish loading !!!")
        except(Exception):
            print(Exception)