import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.kernel_ridge import KernelRidge
from sklearn.linear_model import RidgeCV
from sklearn.model_selection import GridSearchCV, train_test_split
from sklearn.metrics import mean_squared_error

# -------------------------------------------------------
#               MODEL DEFINITIONS
# -------------------------------------------------------

def generate_model_signal(model, n=1000, sigma=0.1, mu=None, seed=111):
    rng = np.random.RandomState(seed)
    d = rng.rand(n) * sigma
    e = rng.randn(n)

    if model == 1:
        # Linear AR(2)
        for i in range(10, n):
            d[i] = 0.7*d[i-1] - 0.8*d[i-2] + e[i]*sigma

    elif model == 2:
        # Nonlinear AR(2)
        for i in range(10, n):
            d[i] = (0.7*d[i-1] - 0.8*d[i-2]
                    -0.7*d[i-1]**2 -0.6*d[i-2]**2 +0.5*d[i-1]*d[i-2]
                    + e[i]*sigma)

    elif model == 3:
        # Logistic map
        if mu is None:
            raise ValueError("Model 3 requires μ.")
        for i in range(10, n):
            d[i] = mu * d[i-1] * (1 - d[i-1])

    return d, e


# -------------------------------------------------------
#       BUILD AR DESIGN MATRIX FOR GIVEN p AND k=1
# -------------------------------------------------------

def build_AR_matrix(x, p):
    n = len(x)
    k = 1
    X = np.zeros((n-p-k+1, p))
    ind = np.arange(p)
    for i in range(n-p-k+1):
        X[i,:] = x[ind]
        ind += 1
    y = x[p+k-1 : n]
    return X, y


# -------------------------------------------------------
# Hyperparameter grid
# -------------------------------------------------------

alpha_grid = np.logspace(-6, -1, 6)
gamma_grid = np.logspace(-9, 3, 13)

param_grid = {"alpha": alpha_grid, "gamma": gamma_grid}

kernels = ["rbf", "laplacian", "poly"]
p_values = [1, 2, 5]
mu_values = [1, 2, 4]   # used only for model 3


# -------------------------------------------------------
# STORAGE
# -------------------------------------------------------

results = []


# -------------------------------------------------------
# MAIN LOOP
# -------------------------------------------------------

for model_idx in [1, 2, 3]:      # linear, nonlinear, logistic
    for p in p_values:
        for mu in (mu_values if model_idx == 3 else [None]):

            # Display progress -------------------------------------------------
            print(f"\n=== Running Model {model_idx}, p={p}, mu={mu} ===")

            # Generate signal
            d, e = generate_model_signal(model_idx, mu=mu)

            # Build AR regressors
            X, y = build_AR_matrix(d, p)

            # Train-test split
            X_train, X_test, y_train, y_test = train_test_split(
                X, y, test_size=0.20, shuffle=False, random_state=2
            )

            # Linear ridge baseline (for reference)
            lin = RidgeCV(alphas=alpha_grid, cv=5)
            lin.fit(X_train, y_train)
            y_pred_lin = lin.predict(X_test)
            rmse_test_lin = np.sqrt(mean_squared_error(y_test, y_pred_lin))
            nrmse_test_lin = rmse_test_lin / 0.1  # sigma = 0.1

            # Loop over kernels
            for ker in kernels:
                print(f"   -> Kernel = {ker} ... running grid search")

                clf = GridSearchCV(
                    KernelRidge(kernel=ker),
                    param_grid=param_grid,
                    cv=5,
                    return_train_score=False
                )

                clf.fit(X_train, y_train)

                # Best model predictions
                y_train_hat = clf.predict(X_train)
                y_test_hat = clf.predict(X_test)

                # RMSE
                rmse_train = np.sqrt(mean_squared_error(y_train, y_train_hat))
                rmse_test  = np.sqrt(mean_squared_error(y_test,  y_test_hat))

                # Normalized RMSE
                nrmse_train = rmse_train / 0.1
                nrmse_test  = rmse_test  / 0.1

                # Store
                results.append({
                    "Model": ["Model 1 (Linear AR)",
                              "Model 2 (Nonlinear AR)",
                              "Model 3 (Logistic)"][model_idx-1],
                    "p": p,
                    "mu": mu,
                    "Kernel": ker,
                    "Best α": clf.best_params_['alpha'],
                    "Best γ": clf.best_params_['gamma'],
                    "KRR Train NRMSE": nrmse_train,
                    "KRR Test NRMSE": nrmse_test,
                    "Linear Test NRMSE": nrmse_test_lin
                })


# -------------------------------------------------------
# FINAL TABLE
# -------------------------------------------------------

df = pd.DataFrame(results)
print("\n\n=== FINAL RESULTS TABLE (Normalized RMSE) ===\n")
print(df)

# Save table as CSV if needed
df.to_csv("final_results_nrmse.csv", index=False)
